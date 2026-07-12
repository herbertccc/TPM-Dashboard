#!/usr/bin/env python3
"""在工作流中把 .b64 文件解码回 .xlsx"""
import base64, glob, os, sys

for f in glob.glob('**/*.b64', recursive=True):
    out = f[:-4]
    with open(out, 'wb') as fout:
        fout.write(base64.b64decode(open(f).read()))
    print(f'  {f} -> {out} ({os.path.getsize(out)} bytes)')
print('Done.')
