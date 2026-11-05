#!/usr/bin/env python3
"""把儀表板需要的檔案複製到 _site/，發布時只上傳這個資料夾。

採白名單：新增前台檔案時要加進 FILES，其他檔案一律不會被發布。
"""
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / '_site'
FILES = [
    'index.html',
    'converter.html',
    'data.json',
    'chart.umd.js',
    'xlsx.full.min.js',
    '_headers',
    'apple-touch-icon.png',
    'favicon-16x16.png',
    'favicon-32x32.png',
    'favicon.ico',
    'icon-192.png',
    'icon-512.png',
]


def main():
    missing = [f for f in FILES if not (ROOT / f).is_file()]
    if missing:
        print('缺少檔案：' + '、'.join(missing), file=sys.stderr)
        return 1
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir()
    for f in FILES:
        shutil.copy2(ROOT / f, OUT / f)
    print(f'{len(FILES)} 個檔案 -> {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
