#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""index.html 自動化驗證腳本。

修改 HTML / JS / CSS 後必跑，全部通過才可輸出／commit。
用法：  python3 scripts/validate_html.py [index.html]
離開碼： 0 = 全部通過、1 = 有項目失敗
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HTML = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "index.html"

results = []      # (pass: bool, name: str, detail: str)


def check(name, ok, detail=""):
    results.append((bool(ok), name, detail))


def extract_script_blocks(src):
    """回傳 [(起始行號, 內容)]；只取無 src 的內嵌 script。"""
    blocks = []
    for m in re.finditer(r"<script(?P<attrs>[^>]*)>(?P<body>.*?)</script>", src, re.S):
        if re.search(r"\bsrc\s*=", m.group("attrs")):
            continue
        line = src[: m.start()].count("\n") + 1
        blocks.append((line, m.group("body")))
    return blocks


def main():
    src = HTML.read_text(encoding="utf-8")

    # ---- 1. Script block 結構正確（開關標籤對齊） --------------------------
    # document.write('<scr'+'ipt ...') 的拆字寫法不會match到真正的標籤，無須排除
    n_open = len(re.findall(r"<script\b", src, re.I))
    n_close = len(re.findall(r"</script\s*>", src, re.I))
    check("1. Script block 開關標籤對齊",
          n_open == n_close,
          f"<script> {n_open} 個 / </script> {n_close} 個")

    blocks = extract_script_blocks(src)
    check("1b. 內嵌 script 區塊可解析", len(blocks) > 0, f"{len(blocks)} 個內嵌區塊")

    # ---- JS 語法檢查（node --check） ---------------------------------------
    node = shutil.which("node")
    if node:
        syntax_ok, syntax_detail = True, []
        for line, body in blocks:
            with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False,
                                             encoding="utf-8") as f:
                f.write(body)
                tmp = f.name
            p = subprocess.run([node, "--check", tmp],
                               capture_output=True, text=True)
            Path(tmp).unlink(missing_ok=True)
            if p.returncode != 0:
                syntax_ok = False
                err = (p.stderr.strip().splitlines() or ["?"])[:4]
                syntax_detail.append(f"第 {line} 行起的區塊：" + " | ".join(err))
        check("1c. JS 語法檢查 (node --check)", syntax_ok,
              "；".join(syntax_detail) if syntax_detail else "全部區塊通過")
    else:
        check("1c. JS 語法檢查 (node --check)", True, "node 不存在，略過")

    # ---- 2. 關鍵資料變數宣告在 DOMContentLoaded 之前 -----------------------
    KEY_VARS = ["STATE", "BIZ_LIST", "BIZ_LIST_TOTAL", "CHANNEL_LIST", "MONTHS"]
    dcl = src.find("DOMContentLoaded")
    if dcl < 0:
        check("2. 關鍵資料變數宣告在 DOMContentLoaded 之前", True,
              "無 DOMContentLoaded，略過")
    else:
        late = []
        for v in KEY_VARS:
            m = re.search(r"^\s*(?:const|let|var)\s+" + re.escape(v) + r"\b",
                          src, re.M)
            if not m:
                late.append(f"{v}(找不到宣告)")
            elif m.start() > dcl:
                late.append(f"{v}(在 DOMContentLoaded 之後)")
        check("2. 關鍵資料變數宣告在 DOMContentLoaded 之前", not late,
              "；".join(late) if late else "、".join(KEY_VARS) + " 皆在之前")

    # ---- 3. CDN 來源只使用 cdnjs.cloudflare.com ---------------------------
    hosts = sorted(set(re.findall(r"https?://([A-Za-z0-9.\-]+)/", src)))
    ALLOWED_CDN = {"cdnjs.cloudflare.com"}
    # 非 CDN 的自家／文件連結不列入判定
    NON_CDN = {"benesse-lc8t.github.io", "benessetw.atlassian.net",
               "www.benesse.com.tw", "store.benesse.com.tw", "shopee.tw",
               "www.w3.org", "claude.ai"}
    bad = [h for h in hosts if h not in ALLOWED_CDN and h not in NON_CDN]
    check("3. CDN 來源只使用 cdnjs.cloudflare.com", not bad,
          ("違規來源：" + ", ".join(bad)) if bad
          else "外部主機：" + ", ".join(hosts))

    # ---- 4. 無重複的 let / const 宣告（同一 script 區塊的頂層） -------------
    dup_detail = []
    for line, body in blocks:
        # 只檢查頂層（無縮排）宣告；區塊內重複宣告由 node --check 攔截
        names = re.findall(r"^(?:let|const)\s+([A-Za-z_$][\w$]*)\s*[=;]",
                           body, re.M)
        seen, dups = set(), set()
        for n in names:
            if n in seen:
                dups.add(n)
            seen.add(n)
        if dups:
            dup_detail.append(f"第 {line} 行起的區塊：{', '.join(sorted(dups))}")
    check("4. 無重複的 let/const 頂層宣告", not dup_detail,
          "；".join(dup_detail) if dup_detail else "無重複")

    # ---- 5. 所有渲染函式都有被呼叫 ----------------------------------------
    js = "\n".join(b for _, b in blocks)
    defined = set(re.findall(r"^\s*function\s+(render[A-Za-z0-9_]*)\s*\(", js, re.M))
    uncalled = []
    for fn in sorted(defined):
        # 宣告以外的出現次數
        uses = len(re.findall(r"\b" + re.escape(fn) + r"\s*\(", js))
        decls = len(re.findall(r"function\s+" + re.escape(fn) + r"\s*\(", js))
        refs = len(re.findall(r"\b" + re.escape(fn) + r"\b", js)) - decls - (uses - decls)
        if uses - decls <= 0 and refs <= 0:
            uncalled.append(fn)
    check("5. 所有 render* 函式都有被呼叫", not uncalled,
          ("未被呼叫：" + ", ".join(uncalled)) if uncalled
          else f"{len(defined)} 個 render 函式全部有呼叫點")

    # ---- 6. getElementById / querySelector 的 id 在 DOM 中存在 -------------
    dom_ids = set(re.findall(r'\bid="([^"]+)"', src))
    # JS 內動態產生的 HTML 字串也算（id="..." 或 id=\"...\" 或 id='...'）
    dom_ids |= set(re.findall(r"\bid='([^']+)'", src))
    dom_ids |= set(re.findall(r"id=\\?\"([^\"\\]+)\\?\"", js))
    dom_ids |= set(re.findall(r"id=([A-Za-z0-9_\-]+)", js))
    missing = []
    for m in re.finditer(r'getElementById\(\s*"([^"]+)"\s*\)', js):
        if m.group(1) not in dom_ids:
            missing.append(m.group(1))
    for m in re.finditer(r"getElementById\(\s*'([^']+)'\s*\)", js):
        if m.group(1) not in dom_ids:
            missing.append(m.group(1))
    for m in re.finditer(r'querySelector(?:All)?\(\s*[\'"]#([A-Za-z0-9_\-]+)[\'"]', js):
        if m.group(1) not in dom_ids:
            missing.append(m.group(1))
    missing = sorted(set(missing))
    check("6. 無殘留的 DOM 元素引用", not missing,
          ("找不到對應元素：" + ", ".join(missing)) if missing
          else "所有靜態 id 引用皆存在")

    # ---- 7. 新功能格式檢查 -------------------------------------------------
    fmt = []
    # 7a. 數字字型：不得再出現等寬字型（斜線零來源）
    mono = re.findall(r"font-family:[^;{}]*(?:monospace|Monaco|Consolas|Courier)[^;{}]*",
                      src, re.I)
    if mono:
        fmt.append(f"仍有等寬字型宣告 {len(mono)} 處")
    # 7b. 數字字型變數已定義且被套用
    if "--font-num" not in src:
        fmt.append("未定義 --font-num")
    if 'font-feature-settings' not in src:
        fmt.append("未設定 font-feature-settings（需關閉斜線零 zero 0）")
    elif not re.search(r'"zero"\s*0', src):
        fmt.append('font-feature-settings 未含 "zero" 0（斜線零未關閉）')
    # 7c. 月末預估區塊：單位「萬」與倍率「x」的顯示
    if "forecast" in src:
        if "fcst-" not in src:
            fmt.append("月末預估區塊缺少 fcst- 前綴元素")
        for need, label in [("月末預估", "標題文字"), ("倍率", "倍率欄位")]:
            if need not in src:
                fmt.append(f"月末預估區塊缺少{label}「{need}」")
    check("7. 新功能格式正確套用", not fmt,
          "；".join(fmt) if fmt else "字型、單位、倍率格式皆正確")

    # ---- 8. 本機 fallback 檔案存在 ----------------------------------------
    local = []
    for m in re.finditer(r'<script[^>]+src="(?!https?:)([^"]+)"', src):
        if not (HTML.parent / m.group(1)).exists():
            local.append(m.group(1))
    for m in re.finditer(r"'<scr'\s*\+\s*'ipt src=\"(?!https?:)([^\"]+)\"", src):
        if not (HTML.parent / m.group(1)).exists():
            local.append(m.group(1))
    check("8. 本機 script fallback 檔案存在", not local,
          ("缺少：" + ", ".join(local)) if local else "全部存在")

    # ---- 9. data.json 可解析且欄位齊全 ------------------------------------
    dj = HTML.parent / "data.json"
    if dj.exists():
        try:
            d = json.loads(dj.read_text(encoding="utf-8"))
            need = ["budget", "actuals", "actualRows", "master",
                    "workingDays", "holidays"]
            lack = [k for k in need if k not in d]
            check("9. data.json 結構完整", not lack,
                  ("缺少 key：" + ", ".join(lack)) if lack
                  else f"actualRows {len(d['actualRows'])} 筆")
        except Exception as e:
            check("9. data.json 結構完整", False, str(e))
    else:
        check("9. data.json 結構完整", True, "檔案不存在，略過")

    # ---- 輸出 --------------------------------------------------------------
    print("=" * 74)
    print(f"index.html 驗證：{HTML}")
    print("=" * 74)
    failed = 0
    for ok, name, detail in results:
        mark = "PASS" if ok else "FAIL"
        if not ok:
            failed += 1
        print(f"[{mark}] {name}")
        if detail:
            print(f"       {detail}")
    print("-" * 74)
    if failed:
        print(f"*** {failed} 項失敗 —— 修正後重跑，全部通過才可輸出 ***")
        return 1
    print(f"全部 {len(results)} 項通過")
    return 0


if __name__ == "__main__":
    sys.exit(main())
