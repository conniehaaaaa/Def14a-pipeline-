#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SEC DEF 14A 質押分析 Pipeline（框架版）
==========================================

依老師交辦的流程（步驟 0-4）重組。延續 v7 爬蟲的核心（鎖定受益所有權表、
SHROWN_A/SHROWN_B、三種 holder 類型、斷點續跑），但把任務切成可分工的階段：

  步驟 0  掃描清單上的 DEF 14A filings（抓原始檔案）
  步驟 1  用寬鬆關鍵字鎖定「受益所有權」章節的「表格 + 附註」，截取出來
  步驟 2  【無質押公司】把表格餵 AI → 取得所有內部人持股   ← AI 問句待老師提供
  步驟 3  【有質押公司】把附註餵 AI → 誰質押/押多少/該人持股 ← AI 問句待老師提供
  步驟 4  【有質押內部人】查其在公司職位（AI 或爬蟲皆可）   ← 待老師提供

現在這版：步驟 0、1 可運作；步驟 2-4 為接好的 stub，等老師的問句填入即可。

pip install requests beautifulsoup4 lxml pandas openpyxl
"""

import os
import re
import time
import json
import html
import pandas as pd
import requests
from bs4 import BeautifulSoup


# ============================================================
# 設定
# ============================================================
INPUT_XLSX      = "DEF14A2023_PLEDGE_RA2_202605.xlsx"   # 清單：CIK, FILEDATE, http
OUTPUT_SECTIONS = "step1_sections.jsonl"                # 步驟1 截取的章節（表格+附註）
OUTPUT_OWNERS   = "step2_owners.csv"                    # 步驟2 內部人持股
OUTPUT_PLEDGES  = "step3_pledges.csv"                   # 步驟3 質押明細
OUTPUT_TITLES   = "step4_titles.csv"                    # 步驟4 質押人職位
PROGRESS_JSON   = "pipeline_progress.json"

LIMIT      = 20          # 試跑筆數；正式跑改 None
FRESH      = True        # True = 忽略舊進度檔重跑（不用手動 del）；正式跑改 False
DIAG_CIK   = "0001031316"  # 對這個 CIK 的 filing 做深度診斷（表格=0 的舊格式）；不要診斷設 None
SLEEP      = 0.8         # SEC EDGAR 禮貌間隔
SAVE_EVERY = 50
TIMEOUT    = 30
RETRIES    = 3
USER_AGENT = "WenZhi Research yizhen1426@gmail.com"

HEADERS = {"User-Agent": USER_AGENT,
           "Accept-Encoding": "gzip, deflate",
           "Host": "www.sec.gov"}


# ============================================================
# 關鍵字 / 正則（步驟1 的核心）
# ============================================================

# 受益所有權章節「標題」的寬鬆關鍵字集合
# —— 依老師記事本的建議：不要只比對單一標題，用一組寬鬆關鍵字
OWNERSHIP_TITLE = re.compile(
    r"security ownership"
    r"|beneficial ownership"
    r"|stock ownership"
    r"|share ownership"
    r"|ownership of (?:certain |our )?(?:common stock|shares|securities|management)"
    r"|principal (?:stockholders|shareholders)"
    r"|certain beneficial owners"
    r"|voting securities and principal holders"
    r"|equity ownership"
    r"|who owns our stock"
    r"|shareholdings of",
    re.I,
)

# 質押相關關鍵字（步驟3 分流用：判斷該公司/該附註是否提到質押）
PLEDGE_KW = re.compile(
    r"\bpledg(?:e|ed|es|ing)\b"          # pledge / pledged / pledging
    r"|\bhypothecat"                      # hypothecated
    r"|held in (?:a )?margin account"     # margin account 持有
    r"|margin loan"
    r"|as collateral"
    r"|securing (?:a |an )?(?:loan|indebtedness)"
    r"|pledged as security",
    re.I,
)

# 表格內容特徵：像持股表應同時有 name + shares/beneficial + percent
TABLE_HAS_NAME  = re.compile(r"name|beneficial owner", re.I)
TABLE_HAS_SHARE = re.compile(r"beneficial|shares|number", re.I)
TABLE_HAS_PCT   = re.compile(r"%|percent", re.I)

# 排除明顯非持股表（薪酬/審計/選舉/選擇權）
TABLE_EXCLUDE = re.compile(
    r"salary|bonus|audit fees|fee category|base salary"
    r"|compensation table|grant date|exercise price|option awards",
    re.I,
)

# 數字（股數）：含千分位或 4 位以上
NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})+|\d{4,}")


# ============================================================
# HTTP / 正規化（延續 v7）
# ============================================================
def fetch(url):
    """抓取單一 filing，含重試與 403/429 退避。"""
    for attempt in range(RETRIES):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            if r.status_code == 200:
                r.encoding = r.apparent_encoding or "utf-8"
                return r.text
            if r.status_code in (403, 429):
                time.sleep(SLEEP * (attempt + 2))
                continue
            return None
        except requests.RequestException:
            time.sleep(SLEEP * (attempt + 1))
    return None


def normalize(s):
    """清掉 zero-width space、不間斷空白、彎引號等噪音。"""
    s = html.unescape(s or "")
    s = s.replace("\u200b", "").replace("\xa0", " ").replace("\u00a0", " ")
    s = s.replace("\u2019", "'").replace("\u2018", "'").replace("\u2014", "-")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n[ \t]*\n+", "\n", s)
    return s


def soup_from(htmltxt):
    """建立 soup；舊式 .htm/XML 用 html.parser，其餘 lxml。把 <sup>角標轉 (n)。"""
    head = htmltxt[:500].lower()
    is_xmlish = ("<?xml" in head) or ("xmlns" in head and "<html" not in head)
    parser = "html.parser" if is_xmlish else "lxml"
    soup = BeautifulSoup(htmltxt, parser)
    for sup in soup.find_all("sup"):
        t = sup.get_text(strip=True)
        if re.fullmatch(r"\d{1,3}", t):
            sup.replace_with(f"({t})")
    return soup


# ============================================================
# 步驟 0：掃描清單上的 DEF 14A filings
# ============================================================
def step0_load_list(path=INPUT_XLSX, limit=LIMIT):
    """讀入清單（CIK, FILEDATE, http），回傳 list[dict]。"""
    df = pd.read_excel(path, dtype=str).fillna("")
    # 容錯：欄名大小寫/別名
    cols = {c.lower().strip(): c for c in df.columns}
    cik_c  = cols.get("cik")
    date_c = cols.get("filedate") or cols.get("file_date") or cols.get("date")
    url_c  = cols.get("http") or cols.get("url") or cols.get("link")
    records = []
    for _, row in df.iterrows():
        records.append({
            "CIK":      str(row[cik_c]).strip(),
            "FILEDATE": str(row[date_c]).strip(),
            "http":     str(row[url_c]).strip(),
        })
    if limit:
        records = records[:limit]
    return records


# ============================================================
# 步驟 1：鎖定受益所有權章節，截取「表格 + 附註」
# ============================================================
def is_ownership_table(table):
    """以內容特徵判斷某 <table> 是否為受益所有權表。"""
    tt = table.get_text(" ", strip=True)
    if not (TABLE_HAS_NAME.search(tt) and TABLE_HAS_SHARE.search(tt)
            and TABLE_HAS_PCT.search(tt)):
        return False
    if TABLE_EXCLUDE.search(tt):
        return False
    return True


def find_section_heading(soup):
    """
    找出受益所有權章節的標題節點（標題關鍵字命中）。
    回傳命中的元素清單，供後續定位「表格 + 附註」。
    """
    hits = []
    # 常見標題標籤：h1-h6、b、strong、p、font、span、div、td
    for el in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6",
                             "b", "strong", "p", "font", "span", "div", "td"]):
        txt = el.get_text(" ", strip=True)
        if not txt or len(txt) > 160:        # 標題不會太長
            continue
        if OWNERSHIP_TITLE.search(txt):
            hits.append(el)
    return hits


def extract_footnotes_after(table):
    """
    截取緊接在表格之後的附註區塊（footnotes）。
    附註常見形式：(1) ... / (a) ... / * ... ，散落在後續 <p>/<div>/<td>。
    策略：從 table 往後掃同層/相鄰節點，收集到下一個區段標題或下一張表為止。
    """
    chunks = []
    node = table
    steps = 0
    while node is not None and steps < 40:
        node = node.find_next_sibling() or node.find_next()
        if node is None:
            break
        steps += 1
        if getattr(node, "name", None) == "table":
            break                              # 遇到下一張表停
        txt = normalize(node.get_text(" ", strip=True)) if hasattr(node, "get_text") else ""
        if not txt:
            continue
        # 命中新區段標題就停
        if OWNERSHIP_TITLE.search(txt) and len(txt) < 160:
            break
        # 收集看起來像附註的段落（以 (n)/(a)/* 起頭，或含 pledge 字樣）
        if re.match(r"^\(?\s*[\d a-z\*]{1,3}\s*\)?\s", txt) or PLEDGE_KW.search(txt):
            chunks.append(txt)
    return "\n".join(chunks)


def step1_extract_section(record, htmltxt):
    """
    回傳 dict：含截取到的受益所有權表（HTML+純文字）、附註、是否含質押。
    這份輸出就是要餵給步驟 2/3 AI 的原料。
    """
    soup = soup_from(htmltxt)
    headings = find_section_heading(soup)

    ownership_tables = []
    for t in soup.find_all("table"):
        if is_ownership_table(t):
            ownership_tables.append(t)

    # 截取表格文字 + 緊接附註
    sections = []
    for t in ownership_tables:
        table_text = normalize(t.get_text("\n", strip=True))
        footnotes  = extract_footnotes_after(t)
        sections.append({
            "table_text": table_text,
            "table_html": str(t),
            "footnotes":  footnotes,
        })

    combined = " ".join(s["table_text"] + " " + s["footnotes"] for s in sections)
    has_pledge = bool(PLEDGE_KW.search(combined))

    return {
        "CIK":          record["CIK"],
        "FILEDATE":     record["FILEDATE"],
        "http":         record["http"],
        "n_headings":   len(headings),
        "n_tables":     len(sections),
        "sections":     sections,
        "has_pledge":   has_pledge,      # ← 步驟2/3 分流依據
    }


# ============================================================
# 步驟 2：【無質押公司】AI → 所有內部人持股
#   AI 問句待老師提供。接口已備好。
# ============================================================
def step2_owners_no_pledge(section):
    """
    輸入：step1 的 section dict（has_pledge == False）
    輸出：list[dict] 內部人持股
    TODO(老師問句)：把 section['sections'] 的 table_text 餵 AI，
                    取回每位內部人 {FULLNAME, SHROWN_A, SHROWN_B, ...}
    """
    raise NotImplementedError("步驟2 AI 問句待老師提供")


# ============================================================
# 步驟 3：【有質押公司】AI → 質押明細
#   AI 問句待老師提供。接口已備好。
# ============================================================
def step3_pledges(section):
    """
    輸入：step1 的 section dict（has_pledge == True）
    輸出：list[dict] 質押明細
    TODO(老師問句)：把 section['sections'] 的 footnotes（+ table_text）餵 AI，
                    取回 {誰質押 PLEDGOR, 質押多少 PLEDGED_SHARES,
                          該人持股 SHROWN, ...}
    """
    raise NotImplementedError("步驟3 AI 問句待老師提供")


# ============================================================
# 步驟 4：【有質押內部人】查其在公司職位
#   AI 或爬蟲皆可。接口已備好。
# ============================================================
def step4_lookup_title(cik, fullname, htmltxt=None):
    """
    輸入：CIK + 質押人姓名（可帶該 filing 全文）
    輸出：該人在公司職位字串
    TODO(老師)：方案A 問 AI 掃 DEF 14A；方案B 從全文比對姓名鄰近的 title 字樣。
    """
    raise NotImplementedError("步驟4 待老師提供（AI 問句 或 爬蟲規則）")


# ============================================================
# 診斷：拆解 表格=0 的舊格式 filing
# ============================================================
def diagnose(record, htmltxt):
    """印出舊格式 filing 的結構，判斷為何抓不到表格。"""
    print("\n" + "=" * 60)
    print(f"診斷 CIK={record['CIK']}  FILEDATE={record['FILEDATE']}")
    print(f"URL: {record['http']}")
    print("=" * 60)

    head = htmltxt[:300].lower()
    is_xmlish = ("<?xml" in head) or ("xmlns" in head and "<html" not in head)
    print(f"檔案前 300 字判斷：{'XML/舊式' if is_xmlish else 'HTML'}")
    print(f"全文長度：{len(htmltxt)} 字")

    soup = soup_from(htmltxt)

    # 1) 有幾個 <table>？
    tables = soup.find_all("table")
    print(f"\n<table> 標籤數：{len(tables)}")

    # 2) 是否為 <pre> 純文字排版的舊 filing？
    pres = soup.find_all("pre")
    print(f"<pre> 標籤數：{len(pres)}（若 >0 多半是純文字排版，表格不在 <table> 裡）")

    # 3) 全文是否出現受益所有權標題關鍵字？出現在哪
    full = normalize(soup.get_text("\n"))
    m = OWNERSHIP_TITLE.search(full)
    if m:
        pos = m.start()
        print(f"\n標題關鍵字命中：'{full[pos:pos+60].strip()}'")
        print("命中處前後文：")
        print("  " + full[max(0, pos - 80):pos + 200].replace("\n", " ⏎ "))
    else:
        print("\n⚠ 全文找不到任何受益所有權標題關鍵字（標題寫法可能是新變體）")

    # 4) 是否出現質押字樣
    pm = PLEDGE_KW.search(full)
    print(f"\n質押關鍵字：{'命中 -> ' + repr(full[pm.start():pm.start()+50]) if pm else '無'}")

    # 5) 若有 <table>，逐張看它為何沒被判定為 ownership table
    if tables:
        print(f"\n逐張檢查 {len(tables)} 張 table：")
        for idx, t in enumerate(tables):
            tt = t.get_text(" ", strip=True)
            checks = (
                f"name={'Y' if TABLE_HAS_NAME.search(tt) else 'N'} "
                f"share={'Y' if TABLE_HAS_SHARE.search(tt) else 'N'} "
                f"pct={'Y' if TABLE_HAS_PCT.search(tt) else 'N'} "
                f"excluded={'Y' if TABLE_EXCLUDE.search(tt) else 'N'}"
            )
            print(f"  [table {idx}] {checks}  前80字: {tt[:80]}")
    print("=" * 60 + "\n")


# ============================================================
# 進度（斷點續跑）
# ============================================================
def load_progress():
    if os.path.exists(PROGRESS_JSON):
        with open(PROGRESS_JSON, encoding="utf-8") as f:
            return set(json.load(f).get("done", []))
    return set()


def save_progress(done):
    with open(PROGRESS_JSON, "w", encoding="utf-8") as f:
        json.dump({"done": sorted(done)}, f)


# ============================================================
# 主流程
# ============================================================
def main():
    records = step0_load_list()
    print(f"[步驟0] 載入 {len(records)} 筆 filing")

    done = set() if FRESH else load_progress()
    if FRESH:
        print("[FRESH] 忽略舊進度檔，全部重跑")
    no_pledge, pledged = [], []

    with open(OUTPUT_SECTIONS, "w" if FRESH else "a", encoding="utf-8") as fout:
        for i, rec in enumerate(records, 1):
            key = f"{rec['CIK']}|{rec['FILEDATE']}"
            if key in done:
                continue

            htmltxt = fetch(rec["http"])
            if not htmltxt:
                print(f"  [{i}] 抓取失敗 CIK={rec['CIK']}")
                done.add(key)
                continue

            # 步驟1：截取受益所有權表 + 附註
            section = step1_extract_section(rec, htmltxt)

            # 診斷：對指定 CIK 中表格=0 的 filing 做深度拆解
            if DIAG_CIK and rec["CIK"] == DIAG_CIK and section["n_tables"] == 0:
                diagnose(rec, htmltxt)

            fout.write(json.dumps(section, ensure_ascii=False) + "\n")

            tag = "有質押" if section["has_pledge"] else "無質押"
            print(f"  [{i}] CIK={rec['CIK']} 表格={section['n_tables']} {tag}")

            # 分流（步驟2/3 待老師問句，先只分類）
            (pledged if section["has_pledge"] else no_pledge).append(section)

            done.add(key)
            if i % SAVE_EVERY == 0:
                save_progress(done)
            time.sleep(SLEEP)

    save_progress(done)
    print(f"\n[彙總] 無質押 {len(no_pledge)} 家、有質押 {len(pledged)} 家")
    print("步驟1 截取結果已寫入：", OUTPUT_SECTIONS)
    print("步驟2-4 待老師提供 AI 問句後接上即可。")


if __name__ == "__main__":
    main()