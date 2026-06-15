#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""슬라이드 겹침 수정 패치 v2"""

path = 'scripts/make_academic_ppt.py'
content = open(path, encoding='utf-8').read()

# 각 교체 대상을 찾아서 보여주기
targets = [
    ('l=6.55, t=1.12, w=6.53', 'l=6.60, t=1.12, w=6.50'),
    ('l=8.08, t=1.12, h=2.0', 'l=7.88, t=1.12, h=2.0'),
    ('1.5, 1.9, 1.9, 1.15, 1.55, 1.35, 1.9, 2.0', '1.35, 1.8, 1.8, 1.05, 1.45, 1.2, 1.8, 1.85'),
    ('h=3.35,', 'h=3.05,'),
]

for old, new in targets:
    if old in content:
        content = content.replace(old, new)
        print(f"OK  : {old[:50]}")
    else:
        print(f"MISS: {old[:50]}")

open(path, 'w', encoding='utf-8').write(content)
print("저장 완료")
