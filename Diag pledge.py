#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""找出真正持股表的位置與它前面的標題，看為何沒被認成候選。"""
import re, html, requests
from bs4 import BeautifulSoup
URL = "http://www.sec.gov/Archives/edgar/data/1032033/000095013307001622/w29674def14a.htm"
UA  = "WenZhi Research yizhen1426@gmail.com"
def normalize(s):
    s = html.unescape(s or "")
    for a,b in [("\u200b",""),("\xa0"," "),("\u00a0"," "),("\u2019","'"),("\u2018","'"),("\u2014","-")]:
        s=s.replace(a,b)
    return re.sub(r"[ \t]+"," ",s)
r=requests.get(URL,headers={"User-Agent":UA},timeout=30)
r.encoding=r.apparent_encoding or "utf-8"
full=normalize(BeautifulSoup(r.text,"lxml").get_text("\n"))
lines=[ln.strip() for ln in full.split("\n")]

# 找「Beneficially Owned」表頭出現處（真持股表標誌）
HDR=re.compile(r"beneficially owned|amount and nature of beneficial",re.I)
hits=[i for i,ln in enumerate(lines) if HDR.search(ln)]
print(f"'Beneficially Owned'類表頭出現在行: {hits}\n")
for h in hits[:4]:
    print(f"===== 表頭行 {h}: {lines[h][:80]!r} =====")
    print("  前 15 行（找標題）:")
    for j in range(max(0,h-15), h):
        if lines[j].strip():
            print(f"    行{j}: {lines[j][:75]!r}")
    print()