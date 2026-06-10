#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
質押命中診斷：印出每筆 filing 裡 PLEDGE_KW 實際命中的句子前後文，
用來判斷是「真質押」還是被反避險/反質押政策（prohibit pledging）誤抓。
不動主程式。跑：python diag_pledge.py
"""
import re, time, html, requests, pandas as pd
from bs4 import BeautifulSoup

INPUT_XLSX = "DEF14A2023_PLEDGE_RA2_202605.xlsx"
LIMIT      = 20
USER_AGENT = "WenZhi Research yizhen1426@gmail.com"
HEADERS = {"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate", "Host": "www.sec.gov"}

PLEDGE_KW = re.compile(
    r"\bpledg(?:e|ed|es|ing)\b|\bhypothecat|held in (?:a )?margin account"
    r"|margin loan|as collateral|securing (?:a |an )?(?:loan|indebtedness)"
    r"|pledged as security", re.I)

# 反向陳述偵測：句中若含禁止/不得 + pledge，多半是反質押政策（偽陽性）
NEGATION = re.compile(
    r"prohibit|not permitted|may not|are not allowed|forbid|restrict(?:s|ed|ion)?"
    r"|policy (?:against|prohibits)|no (?:director|officer|employee).{0,40}pledg", re.I)


def normalize(s):
    s = html.unescape(s or "")
    for a, b in [("\u200b", ""), ("\xa0", " "), ("\u00a0", " "),
                 ("\u2019", "'"), ("\u2018", "'"), ("\u2014", "-")]:
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s)


def fetch(url):
    for k in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
            if r.status_code == 200:
                r.encoding = r.apparent_encoding or "utf-8"
                return r.text
            time.sleep(1.5 * (k + 1))
        except requests.RequestException:
            time.sleep(1.0 * (k + 1))
    return None


def main():
    df = pd.read_excel(INPUT_XLSX, dtype=str).fillna("")
    cols = {c.lower().strip(): c for c in df.columns}
    cik_c = cols.get("cik"); url_c = cols.get("http") or cols.get("url")
    date_c = cols.get("filedate") or cols.get("date")
    rows = df.head(LIMIT) if LIMIT else df

    for i, row in rows.iterrows():
        cik, url = str(row[cik_c]).strip(), str(row[url_c]).strip()
        date = str(row[date_c]).strip()
        txt = fetch(url)
        if not txt:
            print(f"[{i+1}] CIK={cik} 抓取失敗"); continue
        full = normalize(BeautifulSoup(txt, "lxml").get_text(" "))

        hits = list(PLEDGE_KW.finditer(full))
        print(f"\n[{i+1}] CIK={cik} {date}  命中 {len(hits)} 次")
        for m in hits[:6]:                       # 每筆最多看 6 句
            s = max(0, m.start() - 90); e = min(len(full), m.end() + 90)
            ctx = full[s:e]
            neg = "  ← 疑似反質押政策(偽陽性)" if NEGATION.search(ctx) else ""
            print(f"    …{ctx}…{neg}")
        time.sleep(0.8)


if __name__ == "__main__":
    main()