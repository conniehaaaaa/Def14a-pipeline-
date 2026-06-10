# SEC DEF 14A 質押分析 Pipeline

從 SEC EDGAR 的 DEF 14A（委託書 / proxy statement）中，定位並截取「公司管理階層、主要股東及董事持股」（受益所有權，Item 403）章節，判斷公司是否有內部人質押，並把章節原料整理成可餵給 AI 判讀的格式。

本專案延續先前的 DEF 14A 爬蟲，依指導教授交辦重新組織成 5 個步驟的流程。目前**步驟 0、1 已完成並驗證**；步驟 2–4 為已接好的函式接口（stub），等 AI 問句填入。

---

## 流程概觀

| 步驟 | 內容 | 狀態 |
|------|------|------|
| 步驟 0 | 依清單掃描 DEF 14A filings，抓取原始檔案 | 完成 |
| 步驟 1 | 用關鍵字定位受益所有權章節，截取「表格 + 附註」，並判斷是否含質押 | 完成 |
| 步驟 2 | 【無質押公司】把表格餵 AI → 取得所有內部人持股 | 接口已備，待 AI 問句 |
| 步驟 3 | 【有質押公司】把附註餵 AI → 誰質押 / 押多少 / 該人持股 | 接口已備，待 AI 問句 |
| 步驟 4 | 【有質押內部人】查其在公司職位 | 接口已備，待 AI 問句或爬蟲規則 |

---

## 環境需求

```
pip install requests beautifulsoup4 lxml pandas openpyxl
```

Python 3.8+。

---

## 輸入

清單檔（預設 `DEF14A2023_PLEDGE_RA2_202605.xlsx`），需包含以下欄位（大小寫不拘，支援常見別名）：

| 欄位 | 說明 |
|------|------|
| `CIK` | 公司在 SEC 的識別碼 |
| `FILEDATE` | 該 filing 的申報日期 |
| `http` | 該 filing 文件的 URL |

> 注意：此清單為**預篩過的質押樣本**，因此多數（或全部）公司本就有質押。正式跑完整資料時，步驟 2（無質押那條路）可能很少觸發，主力在步驟 3。同一家公司會跨多個年度出現，所有年度都保留（分析以「公司 × 年度」為單位）。

---

## 執行

```
python "Def14a pipeline.py"
```

主要設定集中在檔案開頭：

| 設定 | 用途 |
|------|------|
| `LIMIT` | 試跑筆數；正式跑改 `None` 跑全部 |
| `FRESH` | `True` 時忽略舊進度檔、整批重跑（不必手動刪檔）；正式跑可改 `False` 以支援斷點續跑 |
| `DIAG_CIK` | 對指定 CIK 中「定位失敗」的 filing 印出深度診斷；不需診斷設 `None` |
| `SLEEP` | 對 SEC EDGAR 的禮貌間隔（秒） |
| `USER_AGENT` | SEC 要求帶可聯絡的 User-Agent，請改成自己的 |

---

## 輸出

### `step1_sections.jsonl`

每行一筆 filing，JSON 格式。重要欄位：

| 欄位 | 說明 |
|------|------|
| `CIK` / `FILEDATE` / `http` | 來源識別 |
| `anchor_type` | 章節定位方式：`text`（成功）／`None`（定位失敗，會觸發診斷） |
| `sections[].table_text` | **章節全文**：持股表資料 + 附註，餵 AI 的主要原料 |
| `sections[].table_html` | 落在章節範圍內的 HTML table（保留結構，供步驟 2 結構化抽取） |
| `sections[].footnotes` | 章節內文字（含 (1)(2)… 編號附註，質押資訊多在此） |
| `has_pledge` | 是否含質押（**對全文判斷**） |
| `pledge_loc` | 質押命中位置：`in_section`／`out_of_section`／`none` |

> `has_pledge` 固定對**全文**判斷，因為質押揭露有時放在「關聯交易」段、不在持股表章節內。漏抓（false negative）比誤抓嚴重——漏標會讓有質押的公司被錯分到步驟 2、永久消失在分析裡；誤抓只是多走一次步驟 3 的 AI 判讀，成本低且步驟 3 會再過濾。

### 其他輸出（步驟 2–4，待實作）

`step2_owners.csv`、`step3_pledges.csv`、`step4_titles.csv` — 接口已在程式中預留。

---

## 步驟 1 的定位策略（重點）

DEF 14A 的受益所有權章節標題寫法非常多變（"Security Ownership of Certain Beneficial Owners and Management"、"Beneficial Ownership of Common Stock"、"Beneficial Ownership of Voting Stock"…），且 HTML 結構在不同年代、不同公司差異極大。最終採用的方法是：

1. **文字流截取**：把全文轉純文字，從受益所有權標題截到下一個大章節標題（如 EXECUTIVE COMPENSATION）為止。只要質押附註落在這個範圍內就一定被涵蓋，不受 HTML DOM 結構影響。
2. **評分選錨點**：同一標題字串常在目錄（TOC）、附註、文件尾段重複出現。對每個候選標題試截一次並評分，**優先選含持股表表頭字樣（"Beneficially Owned"、"Amount and Nature of Beneficial"…）且有多筆股數列**的候選。這能避開目錄條目（截出來是空的）和文件尾段雜燴（長但無表頭）。
3. **保留範圍內的 HTML table**：用文字指紋比對，把落在章節範圍內的 `<table>` 一併保留，供步驟 2 需要欄位結構時使用。

### 已解決的三類結構問題

- **表格被拆散**：部分公司（如 Franklin Street Properties）的受益所有權表被 HTML 拆成十幾張相鄰的小 `<table>`（表頭一張、每條附註各一張），無法用「單張表」判斷。文字流截取繞過此問題。
- **純文字排版**：早年的 `.txt` filing 用空格對齊排版，持股表根本不在 `<table>` 裡。文字流截取同樣適用。
- **錨點誤判**：目錄條目、含標題字串的長附註、文件尾段雜燴都可能讓錨點選錯。評分制 + 表頭驗證解決。

---

## 程式結構

```
步驟 0   step0_load_list()            讀清單
步驟 1   step1_extract_section()      主截取（呼叫下列輔助）
           ├ extract_text_section()        文字流截取 + 評分選錨點
           ├ collect_tables_in_text_range() 保留範圍內 HTML table
           └ is_toc_line() 等             TOC / 標題 / 附註判斷
步驟 2   step2_owners_no_pledge()     stub（待 AI 問句）
步驟 3   step3_pledges()              stub（待 AI 問句）
步驟 4   step4_lookup_title()         stub（待 AI 問句或爬蟲規則）
診斷     diagnose()                   對定位失敗的 filing 印出結構分析
進度     load_progress / save_progress  斷點續跑
```

---

## 診斷與驗證

- 若某筆 `anchor_type` 為 `None`（定位失敗），且其 CIK 等於 `DIAG_CIK`，會自動印出該 filing 的結構分析（`<table>` 數、`<pre>` 數、標題命中位置、各表特徵），協助判斷問題類型。
- 驗證截取品質：檢查 `step1_sections.jsonl` 每筆的 `pledge_loc` 是否為 `in_section`、`table_text` 長度是否正常（數百至上萬字）、是否含多筆股數列。異常短（如數十字）通常代表錨點誤判或章節被提前切斷。

---

## 注意事項

- SEC EDGAR 對請求頻率有禮貌性要求；請保留 `SLEEP` 間隔並使用可聯絡的 `USER_AGENT`。
- 沙盒或受限網路可能無法存取 `sec.gov`；需在可連外的環境執行抓取。