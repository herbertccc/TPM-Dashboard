#!/usr/bin/env python3
"""本地工具：把 .xlsx 文件编码为 .xlsx.b64 以便提交到 Git
用法: python3 encode_data.py
"""

import base64, glob, os

FILES = [
    'data/E1W.xlsx',
    'milestone/milestone.xlsx',
    'Version_Plan/Version_Plan.xlsx',
    'Stakeholder_List/Stakeholder_List.xlsx',
]

for f in FILES:
    if not os.path.exists(f):
        print(f'  SKIP {f} (not found)')
        continue
    raw = open(f, 'rb').read()
    b64 = base64.b64encode(raw).decode('ascii')
    with open(f + '.b64', 'w') as out:
        out.write(b64)
    print(f'  {f}: {len(raw)} bytes -> {f}.b64')

print('\nDone. Now run: git add *.b64 && git commit && git push')
