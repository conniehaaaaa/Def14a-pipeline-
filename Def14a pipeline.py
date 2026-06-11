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

現在這版：步驟 0、1 可運作；步驟 2-4 已接上 AI（OpenAI API）。

pip install requests beautifulsoup4 lxml pandas openpyxl python-dotenv

API key 放在專案根目錄的 .env 檔（不要推上 git）：
  OPENAI_API_KEY=sk-...
並在 .gitignore 加入一行 .env
"""

import os
import re
import time
import json
import html
import warnings
import pandas as pd
import requests
from bs4 import BeautifulSoup

try:
    from bs4 import XMLParsedAsHTMLWarning
    warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
except ImportError:
    pass

# 載入 .env（若有安裝 python-dotenv）；找專案根目錄的 .env
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # 沒裝 dotenv 也能跑，只要環境變數已由其他方式設定


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

# --- AI（步驟 2/3）設定 ---
RUN_AI       = True       # False = 只做步驟0/1（截取章節），不呼叫 AI
AI_TEST_MODE = True     # True = 只跑第 1 筆並印出「送什麼/收什麼」，驗證 prompt 與 schema；
                          #        驗證 OK 後改 False 全自動跑全部
OPENAI_MODEL = "gpt-5"    # 模型名稱（請依實際可用模型調整，如 gpt-4o 等）
OPENAI_API_KEY_ENV = "OPENAI_API_KEY"   # 從環境變數讀 key（勿把 key 寫進程式）
AI_MAX_TOKENS = 8000      # 上限：GPT-5 系列會先用 token 做內部推理，需留足額度避免答案被截斷
AI_TEMPERATURE = 0        # 抽取任務要穩定、可重現，溫度設 0
AI_SLEEP     = 0.5        # 連續呼叫 API 的間隔

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
    """
    讀入清單，以 http（文件 URL）為單位去重，回傳 list[dict]。
    為何用 http 去重：一份 filing 由 URL 唯一識別。清單裡同一份 filing 會
    因「多個內部人各佔一列」而重複出現（同 CIK+FILEDATE），但那是同一份文件，
    只需抓一次、AI 一次回傳所有質押人。用 CIK|FILEDATE 去重則會誤刪
    「同公司同日但不同文件」的情況（實測 3063 vs 3064，有 1 筆差異）。
    """
    df = pd.read_excel(path, dtype=str).fillna("")
    cols = {c.lower().strip(): c for c in df.columns}
    cik_c  = cols.get("cik")
    date_c = cols.get("filedate") or cols.get("file_date") or cols.get("date")
    url_c  = cols.get("http") or cols.get("url") or cols.get("link")

    records = []
    seen_http = set()
    for _, row in df.iterrows():
        http = str(row[url_c]).strip()
        if not http or http in seen_http:
            continue                      # 以 http 去重：同一份 filing 只留一筆
        seen_http.add(http)
        records.append({
            "CIK":      str(row[cik_c]).strip(),
            "FILEDATE": str(row[date_c]).strip(),
            "http":     http,
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
    # 標題可能被換行拆成相鄰多行（如 "Principal" / "Shareholders"），
    # 故同時比對「單行」與「相鄰 2-3 行合併」後是否命中標題。
    def _merged(idx, span):
        return " ".join(lines[idx:idx + span]).strip()

    candidates = []
    for idx, ln in enumerate(lines):
        if not ln:
            continue
        # 嘗試單行、+下一行、+下兩行；取第一個命中標題的合併字串
        hit_text = None
        for span in (1, 2, 3):
            cand_text = _merged(idx, span)
            if len(cand_text) > 90:
                break
            if OWNERSHIP_TITLE.search(cand_text):
                hit_text = cand_text
                break
        if hit_text is None:
            continue
        if re.search(r"reporting compliance|section\s*16", hit_text, re.I):
            continue
        if is_toc_line(hit_text) or (re.search(r"\b\d{1,3}$", hit_text) and len(hit_text) < 70):
            continue
        if re.match(r"^\(?\s*(?:\d{1,3}|[a-z]|\*)\s*\)", hit_text):
            continue
        if hit_text.rstrip().endswith(".") or re.search(
                r"\b(?:set forth|based (?:solely )?on|consists of|includes)\b", hit_text, re.I):
            continue
        candidates.append((idx, hit_text))

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

    # 錨點評分（以「標題」為主、內容為輔）：
    # 受益所有權揭露常分「內部人(董事/高管)表」與「5% 機構大股東表」兩段。
    # 質押是內部人做的，故主依據是「標題是否指向內部人表」——這比從截出內容
    # 猜測可靠（截出範圍可能過大而使內容特徵失真）。
    SHARE_PAT = re.compile(r"\d{1,3}(?:,\d{3})+|\d+\.\d+\s*%|\b\d{1,2}\s*%")
    HDR = re.compile(r"beneficially owned|amount and nature of beneficial"
                     r"|number of shares|percent of class|% of class", re.I)
    PLEDGE_LOCAL = re.compile(r"\bpledg(?:e|ed|es|ing)\b|as collateral|as security", re.I)
    # 標題明確指向內部人持股表 → 強加分
    TITLE_INSIDER = re.compile(
        r"directors? and (?:executive |named )?officers"
        r"|by (?:directors|management)|management ownership"
        r"|security ownership of (?:management|directors)"
        r"|ownership of (?:directors|management)"
        r"|directors,? (?:nominees|executive)", re.I)
    # 標題指向「非持股表」的雜訊（持股政策、投票說明） → 強扣分
    TITLE_NOISE = re.compile(
        r"ownership guidelines|how (?:to vote|your shares)|voting (?:procedures|instruction)"
        r"|share ownership is|recorded", re.I)

    def _score(sec, title):
        if len(sec) < 150:
            return -100
        n_share = len(SHARE_PAT.findall(sec))
        if n_share < 2:                 # 至少要有兩個股數/百分比（多筆持股列）
            return -100
        score = (100 if HDR.search(sec) else 10) + min(n_share, 50)
        # 標題訊號（主依據）
        if TITLE_INSIDER.search(title):
            score += 200               # 標題明說是董事/高管表 → 強烈優先
        if TITLE_NOISE.search(title):
            score -= 200               # 標題是持股政策/投票說明 → 強烈排除
        # 內容輔助：質押字樣出現 → 這段就是我們要的（質押在內部人表）
        if PLEDGE_LOCAL.search(sec):
            score += 80
        # 內容輔助：截出過長（>15000字）多半範圍失控，輕微扣分
        if len(sec) > 15000:
            score -= 40
        return score

    scored = []
    for idx, title in candidates:
        sec = _extract_from(idx)
        sc = _score(sec, title)
        scored.append((sc, len(sec), sec))

    scored.sort(key=lambda x: (x[0], -x[1]), reverse=True)  # 高分優先；同分取較短(較精準)
    if scored and scored[0][0] > 0:
        return scored[0][2], True

    # 全部候選都不像持股表時，回最長的候選（至少不空手）
    best = max((s[2] for s in scored), key=len, default="")
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

    # 信心評估：程式自己能判斷的可靠度訊號，供人工複查篩選。
    confidence, reasons = _assess_confidence(
        section_text, found, pledge_in_full, pledge_in_section)

    return {
        "CIK":          record["CIK"],
        "FILEDATE":     record["FILEDATE"],
        "http":         record["http"],
        "anchor_type":  anchor_type,         # 用哪個錨點定位到的（None=失敗）
        "n_tables":     len(sections),
        "sections":     sections,
        "has_pledge":   pledge_in_full,      # ← 步驟2/3 分流依據（全文判斷）
        "pledge_loc":   pledge_loc,          # in_section / out_of_section / none
        "confidence":   confidence,          # high / medium / low
        "review_reason": "; ".join(reasons), # 低信心原因（供 needs_review）
    }


# 信心評估用的訊號
_INSIDER_TITLE = re.compile(
    r"directors? and (?:executive |named )?officers|by (?:directors|management)"
    r"|management ownership|security ownership of (?:management|directors)"
    r"|ownership of (?:directors|management)", re.I)
_INST_ONLY = re.compile(r"the following institutions|more than (?:five|5)\s*(?:percent|%)", re.I)


def _assess_confidence(section_text, found, pledge_in_full, pledge_in_section):
    """
    回傳 (confidence, reasons)。判斷依據：
      - 章節是否定位到、長度是否正常
      - 是否抓到內部人持股表（vs 只抓到機構大股東表）
      - has_pledge=True 但質押不在截取的章節內（可能餵 AI 的原料缺質押）
    """
    reasons = []
    if not found or not section_text:
        return "low", ["章節未定位"]
    n = len(section_text)
    if n < 300:
        reasons.append(f"章節過短({n}字)")
    if n > 18000:
        reasons.append(f"章節過長({n}字，可能範圍失控)")
    # 抓到內部人表特徵？（質押人必為內部人）
    has_insider = bool(_INSIDER_TITLE.search(section_text)) or \
        bool(re.search(r"all directors and executive officers|as a group", section_text, re.I))
    looks_institutional = bool(_INST_ONLY.search(section_text))
    if not has_insider and looks_institutional:
        reasons.append("疑似只抓到機構大股東表，非內部人表")
    # has_pledge 但質押不在章節原料內 → 餵 AI 的原料可能缺質押
    if pledge_in_full and not pledge_in_section:
        reasons.append("質押在章節外，AI 原料可能漏質押")
    if reasons:
        # 僅長度偏短/偏長算 medium；涉及抓錯表或漏質押算 low
        severe = any(("機構" in r or "漏質押" in r or "未定位" in r) for r in reasons)
        return ("low" if severe else "medium"), reasons
    return "high", []


# ============================================================
# AI 呼叫（步驟 2/3 共用）
# ============================================================
# 老師提供的 prompt（步驟3 是步驟2 的超集：含彙總持股 + 質押明細）
PROMPT_STEP2 = (
    "Identify aggregate insider shareholdings from the filing. "
    "If the company has multiple share classes, report aggregate insider "
    "ownership separately for each class; otherwise, report total aggregate "
    "insider ownership."
)
PROMPT_STEP3 = (
    "From the filing, identify aggregate insider shareholdings. "
    "If the company has multiple share classes, report aggregate insider "
    "ownership separately by share class; otherwise, report total aggregate "
    "insider ownership. Also identify any insiders who pledged shares for "
    "personal loans, and report each insider's name, pledged shares, and "
    "shares owned. If multiple share classes exist, report pledged shares and "
    "shares owned separately by share class for each insider."
)

# 要求 AI 嚴格回傳的 JSON schema（附在 prompt 後，確保輸出可解析）
SCHEMA_INSTRUCTION = """
Return ONLY a valid JSON object, no prose, no markdown fences. Use this exact schema:
{
  "has_multiple_classes": true/false,
  "aggregate_insider_ownership": [
    {"share_class": "Common" or class name, "shares": number or null}
  ],
  "pledges": [
    {"insider_name": "...", "title": "person's role at the company, e.g. President, CFO, Director",
     "share_class": "Common" or class name,
     "pledged_shares": number or null, "shares_owned": number or null}
  ]
}
Rules:
- If single share class, use "Common" as share_class and one entry in aggregate_insider_ownership.
- If no pledges are found, "pledges" must be an empty list [].
- For "title", use the person's role as stated in the filing (director, officer title, etc.); use null if not stated.
- Use plain integers for share counts (no commas). Use null if a value is not stated.
- Do not invent data; only report what the text supports.
"""


def _openai_key():
    key = os.environ.get(OPENAI_API_KEY_ENV, "")
    if not key:
        raise RuntimeError(
            f"找不到 API key：請在專案根目錄建立 .env 檔，內容：\n"
            f"  {OPENAI_API_KEY_ENV}=sk-...\n"
            f"並確認已 pip install python-dotenv。\n"
            f"（或改用環境變數：PowerShell 執行 $env:{OPENAI_API_KEY_ENV}='sk-...'）")
    return key


def call_openai(section_text, base_prompt):
    """
    呼叫 OpenAI Chat Completions，回傳解析後的 dict（依 SCHEMA_INSTRUCTION）。
    失敗回 None。送出的是 step1 截取的章節文字（情境B：只送相關段落，省 token）。
    """
    user_content = (
        f"{base_prompt}\n{SCHEMA_INSTRUCTION}\n\n"
        f"=== FILING OWNERSHIP SECTION ===\n{section_text}"
    )
    payload = {
        "model": OPENAI_MODEL,
        "messages": [
            {"role": "system", "content":
             "You are a precise financial-filings data extractor. "
             "Output strictly valid JSON per the user's schema."},
            {"role": "user", "content": user_content},
        ],
        "max_completion_tokens": AI_MAX_TOKENS,
    }
    # 較舊模型（gpt-4o 等）用 max_tokens 且支援 temperature；
    # 較新模型（gpt-5 系列）用 max_completion_tokens 且 temperature 固定為預設值。
    if not OPENAI_MODEL.startswith(("gpt-5", "o1", "o3", "o4")):
        payload["max_tokens"] = payload.pop("max_completion_tokens")
        payload["temperature"] = AI_TEMPERATURE
    headers = {"Authorization": f"Bearer {_openai_key()}",
               "Content-Type": "application/json"}
    for attempt in range(RETRIES):
        try:
            r = requests.post("https://api.openai.com/v1/chat/completions",
                              headers=headers, json=payload, timeout=60)
            if r.status_code == 200:
                data = r.json()
                choice = data.get("choices", [{}])[0]
                content = choice.get("message", {}).get("content", "")
                finish = choice.get("finish_reason", "")
                if not content or not content.strip():
                    # 空回傳：多半是答案被截斷（length）或內容過濾。印出診斷。
                    usage = data.get("usage", {})
                    print(f"    [AI] 空回傳 finish_reason={finish} usage={usage}")
                    if finish == "length":
                        print("    [AI] → 被 token 上限截斷，請調高 AI_MAX_TOKENS")
                    # length 截斷時重試一次（已調高上限的話通常下次就過）
                    if finish == "length" and attempt < RETRIES - 1:
                        time.sleep(1)
                        continue
                    return None
                return _parse_json(content)
            if r.status_code in (429, 500, 502, 503):
                time.sleep(2 * (attempt + 1))
                continue
            print(f"    [AI] HTTP {r.status_code}: {r.text[:200]}")
            return None
        except requests.RequestException as e:
            print(f"    [AI] 請求錯誤：{e}")
            time.sleep(2 * (attempt + 1))
    return None


def _parse_json(text):
    """從 AI 回傳擷取 JSON（容錯：去除 ```json 圍欄、抓第一個 {...}）。"""
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?|```$", "", t, flags=re.I).strip()
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", t, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    print(f"    [AI] JSON 解析失敗，原始回傳前 200 字：{t[:200]}")
    return None


def _section_text_of(section):
    """取 step1 截到的章節文字（餵 AI 的原料）。"""
    secs = section.get("sections", [])
    if not secs:
        return ""
    return secs[0].get("table_text", "") or secs[0].get("footnotes", "") or ""


# ============================================================
# 步驟 2：【無質押公司】AI → 彙總內部人持股
# ============================================================
def step2_owners_no_pledge(section):
    """
    輸入：step1 的 section dict（has_pledge == False）
    輸出：list[dict] 彙總持股列（每類股一列），含 CIK/FILEDATE 便於關聯。
    """
    text = _section_text_of(section)
    if not text:
        return []
    result = call_openai(text, PROMPT_STEP2)
    if not result:
        return []
    rows = []
    for agg in result.get("aggregate_insider_ownership", []):
        rows.append({
            "CIK":        section["CIK"],
            "FILEDATE":   section["FILEDATE"],
            "share_class": agg.get("share_class"),
            "aggregate_insider_shares": agg.get("shares"),
        })
    return rows


# ============================================================
# 步驟 3：【有質押公司】AI → 彙總持股 + 逐人質押明細
# ============================================================
def step3_pledges(section):
    """
    輸入：step1 的 section dict（has_pledge == True）
    輸出：(aggregate_rows, pledge_rows) 兩個 list[dict]。
      aggregate_rows：彙總持股（每類股一列）
      pledge_rows：逐質押內部人（姓名/類股/質押股數/持股數）
    """
    text = _section_text_of(section)
    if not text:
        return [], []
    result = call_openai(text, PROMPT_STEP3)
    if not result:
        return [], []
    agg_rows, pledge_rows = [], []
    for agg in result.get("aggregate_insider_ownership", []):
        agg_rows.append({
            "CIK":        section["CIK"],
            "FILEDATE":   section["FILEDATE"],
            "share_class": agg.get("share_class"),
            "aggregate_insider_shares": agg.get("shares"),
        })
    for p in result.get("pledges", []):
        pledge_rows.append({
            "CIK":           section["CIK"],
            "FILEDATE":      section["FILEDATE"],
            "insider_name":  p.get("insider_name"),
            "share_class":   p.get("share_class"),
            "pledged_shares": p.get("pledged_shares"),
            "shares_owned":  p.get("shares_owned"),
            "title":         (p.get("title") or "").strip(),  # AI 直接回傳；空則步驟4備援
        })
    return agg_rows, pledge_rows


# ============================================================
# 步驟 4：【有質押內部人】查其在公司職位
#   AI 或爬蟲皆可。接口已備好。
# ============================================================
def step4_lookup_title(cik, fullname, htmltxt=None):
    """
    輸入：CIK + 質押人姓名（可帶該 filing 全文）
    輸出：該人在公司職位字串
    TODO(老師)：方案A 問 AI 掃 DEF 14A；方案B 從全文比對姓名鄰近的 title 字樣。
    目前先以爬蟲方式嘗試：在全文中找姓名鄰近的 director/officer 字樣。
    """
    if not htmltxt or not fullname:
        return ""
    full = normalize(BeautifulSoup(htmltxt, "lxml").get_text(" "))
    # 找姓名出現處附近 120 字，比對常見職稱字樣
    idx = full.find(fullname)
    if idx == -1:
        return ""
    window = full[max(0, idx - 40): idx + 160]
    m = re.search(
        r"(chief executive officer|chief financial officer|president|"
        r"chairman|chair of the board|director|chief operating officer|"
        r"executive vice president|senior vice president|vice president|"
        r"general counsel|treasurer|secretary|founder)", window, re.I)
    return m.group(1) if m else ""


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
def _write_csv(path, rows, fieldnames):
    """寫 CSV（utf-8-sig 讓 Excel 正確顯示）。"""
    import csv
    if not rows:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def main():
    records = step0_load_list()
    print(f"[步驟0] 載入 {len(records)} 筆 filing")

    done = set() if FRESH else load_progress()
    if FRESH:
        print("[FRESH] 忽略舊進度檔，全部重跑")
    no_pledge, pledged = [], []
    sections_all = []
    failed_fetch = []          # 抓取失敗的 filing，供之後重試

    with open(OUTPUT_SECTIONS, "w" if FRESH else "a", encoding="utf-8") as fout:
        for i, rec in enumerate(records, 1):
            key = rec["http"]            # 以 http 為唯一鍵（斷點續跑用）
            if key in done:
                continue

            htmltxt = fetch(rec["http"])
            if not htmltxt:
                print(f"  [{i}] 抓取失敗 CIK={rec['CIK']}")
                failed_fetch.append(rec)
                done.add(key)
                continue

            # 步驟1：截取受益所有權章節（合併相鄰小表 + 附註）
            section = step1_extract_section(rec, htmltxt)

            if DIAG_CIK and rec["CIK"] == DIAG_CIK and section["anchor_type"] is None:
                diagnose(rec, htmltxt)

            fout.write(json.dumps(section, ensure_ascii=False) + "\n")

            tag = "有質押" if section["has_pledge"] else "無質押"
            if section["pledge_loc"] == "out_of_section":
                tag += "(章節外)"
            anc = section["anchor_type"] or "未定位"
            conf = section["confidence"]
            mark = {"high": "", "medium": " ⚠", "low": " ‼"}[conf]
            print(f"  [{i}] CIK={rec['CIK']} 錨點={anc} {tag} 信心={conf}{mark}")

            (pledged if section["has_pledge"] else no_pledge).append(section)
            sections_all.append((section, htmltxt))

            done.add(key)
            if i % SAVE_EVERY == 0:
                save_progress(done)
            time.sleep(SLEEP)

    save_progress(done)

    # 信心統計 + 待複查清單
    conf_counts = {"high": 0, "medium": 0, "low": 0}
    review_rows = []
    for sec, _ in sections_all:
        conf_counts[sec["confidence"]] += 1
        if sec["confidence"] != "high":
            review_rows.append({
                "CIK": sec["CIK"], "FILEDATE": sec["FILEDATE"], "http": sec["http"],
                "confidence": sec["confidence"], "reason": sec["review_reason"],
            })
    for rec in failed_fetch:
        review_rows.append({
            "CIK": rec["CIK"], "FILEDATE": rec["FILEDATE"], "http": rec["http"],
            "confidence": "low", "reason": "抓取失敗（需重試）"})

    print(f"\n[彙總] 無質押 {len(no_pledge)} 家、有質押 {len(pledged)} 家、抓取失敗 {len(failed_fetch)} 筆")
    print(f"[信心] high={conf_counts['high']} medium={conf_counts['medium']} low={conf_counts['low']}")
    if review_rows:
        _write_csv("needs_review.csv", review_rows,
                   ["CIK", "FILEDATE", "http", "confidence", "reason"])
        print(f"[待複查] {len(review_rows)} 筆寫入 needs_review.csv（medium/low/抓取失敗）")
    print("步驟1 截取結果已寫入：", OUTPUT_SECTIONS)

    if not RUN_AI:
        print("RUN_AI=False：略過步驟2-4。")
        return

    # ---- 步驟 2/3/4：呼叫 AI ----
    owners_rows, agg_rows, pledge_rows = [], [], []

    targets = sections_all
    if AI_TEST_MODE:
        # 測試模式：只跑第 1 筆「有質押」的（步驟3 是超集，最能驗證 schema）
        first_pledge = next((x for x in sections_all if x[0]["has_pledge"]), None)
        targets = [first_pledge] if first_pledge else sections_all[:1]
        print(f"\n[AI 測試模式] 只跑 1 筆驗證 prompt 與 schema "
              f"(CIK={targets[0][0]['CIK']} {targets[0][0]['FILEDATE'][:10]})")

    for sec, htmltxt in targets:
        if sec["has_pledge"]:
            text = _section_text_of(sec)
            if AI_TEST_MODE:
                print("\n===== 送給 AI 的內容（前 600 字）=====")
                print(text[:600])
                print("\n===== Prompt =====")
                print(PROMPT_STEP3[:200] + " ...")
            a_rows, p_rows = step3_pledges(sec)
            # 步驟4：AI 已直接回傳職位；僅當 AI 沒給時，才用爬蟲從全文備援
            for pr in p_rows:
                if not pr.get("title"):
                    pr["title"] = step4_lookup_title(pr["CIK"], pr["insider_name"], htmltxt)
            agg_rows += a_rows
            pledge_rows += p_rows
            if AI_TEST_MODE:
                print("\n===== AI 回傳解析後（彙總持股）=====")
                print(json.dumps(a_rows, ensure_ascii=False, indent=2))
                print("===== AI 回傳解析後（質押明細 + 職位）=====")
                print(json.dumps(p_rows, ensure_ascii=False, indent=2))
        else:
            o_rows = step2_owners_no_pledge(sec)
            owners_rows += o_rows
            if AI_TEST_MODE:
                print("\n===== AI 回傳解析後（無質押→彙總持股）=====")
                print(json.dumps(o_rows, ensure_ascii=False, indent=2))
        time.sleep(AI_SLEEP)

    if AI_TEST_MODE:
        print("\n[AI 測試模式] 驗證完成。schema 正確的話，把 AI_TEST_MODE 改 False 全自動跑。")
        return

    # 寫出 CSV
    _write_csv(OUTPUT_OWNERS, owners_rows,
               ["CIK", "FILEDATE", "share_class", "aggregate_insider_shares"])
    _write_csv(OUTPUT_PLEDGES, pledge_rows,
               ["CIK", "FILEDATE", "insider_name", "share_class",
                "pledged_shares", "shares_owned", "title"])
    # 有質押公司的彙總持股另存（與無質押的合併在 owners，或單獨檢視）
    _write_csv("step3_aggregate.csv", agg_rows,
               ["CIK", "FILEDATE", "share_class", "aggregate_insider_shares"])
    print(f"\n步驟2 無質押彙總持股 → {OUTPUT_OWNERS}（{len(owners_rows)} 列）")
    print(f"步驟3 質押明細 → {OUTPUT_PLEDGES}（{len(pledge_rows)} 列）")
    print(f"步驟3 有質押公司彙總持股 → step3_aggregate.csv（{len(agg_rows)} 列）")


if __name__ == "__main__":
    main()