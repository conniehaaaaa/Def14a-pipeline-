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

# 目錄(TOC)行：標題後跟一串點點/底線 + 頁碼，要跳過（否則章節起點會錨在目錄上）
TOC_LINE = re.compile(r"\.{4,}\s*\d{1,3}\s*$|\s\d{1,3}\s*$")

# 章節結束邊界：受益所有權章節之後的下一個大章節標題，掃到就停止合併
STOP_HEADING = re.compile(
    r"^(?:executive compensation|election of directors|compensation discussion"
    r"|summary compensation|section\s*16|delinquent section|audit committee"
    r"|certain relationships|transactions with|equity compensation plan"
    r"|report of (?:the )?(?:audit|compensation)|proposal\s*\d|ratification"
    r"|corporate governance|director compensation|nominees|other matters"
    r"|delinquent\s+(?:section|filings))", re.I)

# 表頭錨點（內容錨點用）：含「Beneficially Owned」字樣的表頭
HEADER_ANCHOR = re.compile(r"beneficially owned|amount and nature of beneficial", re.I)

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
def is_toc_line(txt):
    """判斷是否為目錄(TOC)行：標題後接點點/底線+頁碼。"""
    return bool(TOC_LINE.search(txt))


def is_ownership_table(table):
    """單張表的內容特徵判斷（保留作備援/診斷用，已非主路徑）。"""
    tt = table.get_text(" ", strip=True)
    if not (TABLE_HAS_NAME.search(tt) and TABLE_HAS_SHARE.search(tt)
            and TABLE_HAS_PCT.search(tt)):
        return False
    if TABLE_EXCLUDE.search(tt):
        return False
    return True


def _iter_elements(soup):
    """依文件出現順序走訪可能是標題的元素 + 所有 table。"""
    tags = ["h1", "h2", "h3", "h4", "h5", "h6", "b", "strong",
            "p", "font", "span", "div", "td", "table"]
    return soup.find_all(tags)


def _is_toc_context(el, txt):
    """
    判斷此標題命中是否落在目錄(TOC)區。涵蓋三種 TOC 格式：
      1) 同行點點+頁碼（is_toc_line 已處理）
      2) 標題行後緊接純頁碼（下一個兄弟/文字節點是純數字）
      3) 標題後直接接頁碼數字（如 "...VOTING STOCK 16"）
    """
    if is_toc_line(txt):
        return True
    # 標題文字結尾就是頁碼（短標題 + 結尾 1-3 位數字，且整體很短）
    if re.search(r"\b\d{1,3}$", txt) and len(txt) < 70:
        # 但要排除真正章節標題剛好以數字結尾的罕見情況：TOC 條目通常後面馬上又是另一個 TOC 條目
        return True
    # 下一個文字節點是純頁碼
    try:
        nxt = el.find_next(string=True)
        if nxt and re.fullmatch(r"\s*\d{1,3}\s*", str(nxt)):
            return True
    except Exception:
        pass
    return False


def find_section_start(soup):
    """
    定位受益所有權章節起點。
    錨點A（標題優先）：命中 OWNERSHIP_TITLE 且非 TOC 行的標題元素。
        同一份文件中標題常出現兩次（目錄 + 正文），錨「最後一次」非 TOC 命中，
        因為正文章節永遠在目錄之後 —— 即使 TOC 偵測漏判也能選到正文。
    錨點B（內容備援）：含「Beneficially Owned」字樣的表頭。
    回傳 (anchor_element, anchor_type) 或 (None, None)。
    """
    candidates = []
    for el in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6",
                             "b", "strong", "p", "font", "span", "div", "td"]):
        txt = el.get_text(" ", strip=True)
        if not txt or len(txt) > 120:
            continue
        if not OWNERSHIP_TITLE.search(txt):
            continue
        # 排除 Section 16(a) Beneficial Ownership Reporting Compliance（非持股章節）
        if re.search(r"reporting compliance|section\s*16", txt, re.I):
            continue
        if _is_toc_context(el, txt):
            continue
        candidates.append(el)

    if candidates:
        # 錨最後一次非 TOC 命中（正文那個）
        return candidates[-1], "title"

    # 錨點B：表頭。找第一張含 Beneficially Owned 的 table。
    for t in soup.find_all("table"):
        if HEADER_ANCHOR.search(t.get_text(" ", strip=True)):
            return t, "header"

    return None, None


def collect_section_blocks(anchor):
    """
    從錨點往下，依文件順序收集連續的 table 與其間/其後的附註文字，
    直到掃到下一個大章節標題（STOP_HEADING）或超出步數上限。
    回傳 (tables_html:list, body_text:str)。

    附註處理：受益所有權表後常接 (1)(2)... 編號附註（質押資訊多在此）。
    這些附註可能是獨立 <p> 或獨立小 <table>，必須一併收進來，
    且不可因附註中含 STOP_HEADING 字樣（如 "transactions"）而誤停。
    """
    tables_html = []
    text_chunks = []
    node = anchor
    steps = 0
    collected = 0          # 已收集的文字/表格塊數，用來確認已離開標題本身
    FOOTNOTE = re.compile(r"^\(?\s*(?:\d{1,3}|[a-z]|\*)\s*\)")  # (1) / (a) / *
    while node is not None and steps < 200:
        node = node.find_next()
        if node is None:
            break
        steps += 1
        name = getattr(node, "name", None)

        if name == "table":
            tables_html.append(str(node))
            text_chunks.append(normalize(node.get_text("\n", strip=True)))
            collected += 1
            continue

        # 只看葉節點文字，避免父容器重複計入
        if name in ("p", "div", "span", "font", "li", "td"):
            if node.find(["p", "div", "table", "span", "font", "li"]):
                continue
            txt = normalize(node.get_text(" ", strip=True))
            if not txt:
                continue
            # 附註行（(1)(a)* 起頭）一律收進來，不視為章節結束
            if FOOTNOTE.match(txt):
                text_chunks.append(txt)
                collected += 1
                continue
            # 已收集過內容後，碰到下一個大章節標題就停
            if collected > 0 and len(txt) < 120 and STOP_HEADING.match(txt) \
                    and not is_toc_line(txt):
                break
            text_chunks.append(txt)
            collected += 1

    return tables_html, "\n".join(c for c in text_chunks if c)


def extract_text_section(soup):
    """
    文字流截取（主路徑）：從正文（非 TOC）的受益所有權標題，
    截到下一個大章節標題（STOP_HEADING）為止，回傳整段純文字。
    這是最穩的邊界方案：不管附註是獨立 table、獨立 p、或被分頁符打斷，
    只要落在「標題」與「下一個大章節」之間就一定被涵蓋（質押附註必在內）。
    回傳 (section_text:str, found:bool)。
    """
    full = normalize(soup.get_text("\n"))
    lines = [ln.strip() for ln in full.split("\n")]

    # 收集所有「可能是標題」的候選行（通過 TOC/附註/完整句排除）。
    candidates = []
    for idx, ln in enumerate(lines):
        if not ln or len(ln) > 90:
            continue
        if not OWNERSHIP_TITLE.search(ln):
            continue
        if re.search(r"reporting compliance|section\s*16", ln, re.I):
            continue
        if is_toc_line(ln) or (re.search(r"\b\d{1,3}$", ln) and len(ln) < 70):
            continue
        if re.match(r"^\(?\s*(?:\d{1,3}|[a-z]|\*)\s*\)", ln):
            continue
        if ln.rstrip().endswith(".") or re.search(
                r"\b(?:set forth|based (?:solely )?on|consists of|includes)\b", ln, re.I):
            continue
        candidates.append(idx)

    if not candidates:
        return "", False

    def _extract_from(start):
        out, seen = [], 0
        for ln in lines[start + 1: start + 1 + 600]:
            if not ln:
                continue
            if re.match(r"^\(?\s*(?:\d{1,3}|[a-z]|\*)\s*\)", ln):
                out.append(ln); seen += 1; continue
            if seen > 0 and len(ln) < 120 and STOP_HEADING.match(ln) and not is_toc_line(ln):
                break
            if re.fullmatch(r"\d{1,3}", ln) or re.match(r"TABLE OF CONTENTS", ln, re.I):
                continue
            out.append(ln); seen += 1
        return "\n".join(out)

    # 錨點驗證 + 評分：受益所有權章節的鐵特徵是
    #   (a) 含表頭字樣 "Beneficially Owned" / "Amount and Nature of Beneficial"
    #   (b) 短距內出現多個「股數/百分比」樣式（持股列）
    # 文件尾段（proxy 雜燴）可能夠長且含零星數字，但不會有表頭字樣 → 用表頭區分真假。
    SHARE_PAT = re.compile(r"\d{1,3}(?:,\d{3})+|\d+\.\d+\s*%|\b\d{1,2}\s*%")
    HDR = re.compile(r"beneficially owned|amount and nature of beneficial"
                     r"|number of shares|percent of class|% of class", re.I)

    def _score(sec):
        if len(sec) < 150:
            return -1
        n_share = len(SHARE_PAT.findall(sec))
        has_hdr = bool(HDR.search(sec))
        if n_share < 2:                 # 至少要有兩個股數/百分比（多筆持股列）
            return -1
        # 有表頭 → 高分；無表頭但多股數 → 中分（仍可能是純文字持股表如 .txt）
        return (100 if has_hdr else 10) + min(n_share, 50)

    scored = []
    for start in candidates:
        sec = _extract_from(start)
        sc = _score(sec)
        if sc > 0:
            scored.append((sc, len(sec), sec))

    if scored:
        # 取分數最高者；同分取較長（資訊較完整）
        scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return scored[0][2], True

    # 全部候選都不像持股表時，回最長的候選（至少不空手）
    best = ""
    for start in candidates:
        sec = _extract_from(start)
        if len(sec) > len(best):
            best = sec
    return best, bool(best)


def collect_tables_in_text_range(soup, section_text):
    """
    輔助：保留落在章節文字範圍內的 HTML table（給步驟2結構化用）。
    判準：table 的純文字內容大部分出現在 section_text 裡。
    """
    if not section_text:
        return ""
    kept = []
    sec_norm = re.sub(r"\s+", "", section_text)
    for t in soup.find_all("table"):
        tt = re.sub(r"\s+", "", normalize(t.get_text(" ", strip=True)))
        if len(tt) < 20:
            continue
        # 取 table 前 60 字當指紋，看是否落在章節文字內
        fp = tt[:60]
        if fp and fp in sec_norm:
            kept.append(str(t))
    return "\n".join(kept)


def step1_extract_section(record, htmltxt):
    """
    回傳 dict：含截取到的受益所有權章節（文字流為主 + 範圍內 HTML table）、是否含質押。
    這份輸出就是要餵給步驟 2/3 AI 的原料。
    """
    soup = soup_from(htmltxt)

    # 主路徑：文字流截取章節（保證附註/質押在內）
    section_text, found = extract_text_section(soup)

    sections = []
    anchor_type = None
    if found and section_text:
        anchor_type = "text"
        table_html = collect_tables_in_text_range(soup, section_text)
        sections.append({
            "anchor_type": "text",
            "table_text":  section_text,    # 章節全文（含表格資料與附註）
            "table_html":  table_html,      # 範圍內保留的 HTML table（輔助結構）
            "footnotes":   section_text,
        })

    # has_pledge 固定對全文判斷（與章節定位脫鉤）。
    # 理由：質押揭露常在「關聯交易」段、不在持股表章節內；漏抓比誤抓嚴重。
    full_text = normalize(soup.get_text(" "))
    pledge_in_full = bool(PLEDGE_KW.search(full_text))
    pledge_in_section = bool(PLEDGE_KW.search(section_text)) if section_text else False
    if pledge_in_section:
        pledge_loc = "in_section"
    elif pledge_in_full:
        pledge_loc = "out_of_section"
    else:
        pledge_loc = "none"

    return {
        "CIK":          record["CIK"],
        "FILEDATE":     record["FILEDATE"],
        "http":         record["http"],
        "anchor_type":  anchor_type,         # 用哪個錨點定位到的（None=失敗）
        "n_tables":     len(sections),
        "sections":     sections,
        "has_pledge":   pledge_in_full,      # ← 步驟2/3 分流依據（全文判斷）
        "pledge_loc":   pledge_loc,          # in_section / out_of_section / none
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

            # 步驟1：截取受益所有權章節（合併相鄰小表 + 附註）
            section = step1_extract_section(rec, htmltxt)

            # 診斷：定位失敗（anchor_type 為 None）時才深度拆解
            if DIAG_CIK and rec["CIK"] == DIAG_CIK and section["anchor_type"] is None:
                diagnose(rec, htmltxt)

            fout.write(json.dumps(section, ensure_ascii=False) + "\n")

            tag = "有質押" if section["has_pledge"] else "無質押"
            if section["pledge_loc"] == "out_of_section":
                tag += "(章節外)"
            anc = section["anchor_type"] or "未定位"
            print(f"  [{i}] CIK={rec['CIK']} 錨點={anc} 區塊={section['n_tables']} {tag}")

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