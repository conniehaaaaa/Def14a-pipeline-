# SEC DEF 14A 質押分析 Pipeline

從 SEC EDGAR 的 DEF 14A（委託書 / proxy statement）中，定位並截取「公司管理階層、主要股東及董事持股」（受益所有權，Item 403）章節，判斷公司是否有內部人質押，並透過 AI（OpenAI API）抽取出彙總內部人持股與逐人質押明細。

本專案依指導教授交辦的流程組織成 5 個步驟。目前**步驟 0–4 全部接通並完成端到端驗證**（20 筆測試樣本通過，含多類股與多質押人的複雜情況）。

---

## 流程概觀

| 步驟 | 內容 | 狀態 |
|------|------|------|
| 步驟 0 | 依清單掃描 DEF 14A filings，抓取原始檔案 | ✅ |
| 步驟 1 | 定位受益所有權章節，截取「表格 + 附註」，判斷是否含質押 | ✅ |
| 步驟 2 | 【無質押公司】把章節餵 AI → 彙總內部人持股（多類股分開） | ✅ |
| 步驟 3 | 【有質押公司】把章節餵 AI → 彙總持股 + 逐人質押明細（姓名/質押股數/持股數/職位，多類股分開） | ✅ |
| 步驟 4 | 質押內部人職位：步驟 3 由 AI 一併回傳；AI 未提供時以爬蟲從全文備援 | ✅ |

> 步驟 2 與步驟 3 的 prompt 由指導教授提供，程式中原文照用。步驟 3 是步驟 2 的超集（含彙總持股 + 質押明細）。

---

## 環境需求

```
pip install requests beautifulsoup4 lxml pandas openpyxl python-dotenv
```

Python 3.8+。

---

## 設定 API Key（.env）

API key 放在專案根目錄的 `.env` 檔，**不要推上 git**：

```
OPENAI_API_KEY=sk-your-key-here
```

`.gitignore` 已排除 `.env`。專案另附 `.env.example` 作為範本（可安全推 git）。程式啟動時以 `python-dotenv` 自動載入 `.env`；未安裝 dotenv 時會退回讀取系統環境變數。

---

## 輸入

清單檔（預設 `DEF14A2023_PLEDGE_RA2_202605.xlsx`），需包含以下欄位（大小寫不拘，支援常見別名）：

| 欄位 | 說明 |
|------|------|
| `CIK` | 公司在 SEC 的識別碼 |
| `FILEDATE` | 該 filing 的申報日期 |
| `http` | 該 filing 文件的 URL |

> 此清單為**預篩過的質押樣本**，多數（或全部）公司本就有質押，故實務上幾乎都走步驟 3。同一家公司會跨多個年度出現，所有年度都保留（分析以「公司 × 年度」為單位）。

---

## 執行

```
python "Def14a pipeline.py"
```

主要設定集中在檔案開頭：

| 設定 | 用途 |
|------|------|
| `LIMIT` | 試跑筆數；正式跑改 `None` 跑全部 |
| `FRESH` | `True` 忽略舊進度檔整批重跑；正式跑（資料量大）建議改 `False` 以支援斷點續跑 |
| `RUN_AI` | `False` 時只做步驟 0/1（截取章節），不呼叫 AI |
| `AI_TEST_MODE` | `True` 時只跑 1 筆並印出「送什麼 / 收什麼」，驗證 prompt 與 schema；驗證 OK 後改 `False` 全自動跑 |
| `OPENAI_MODEL` | 模型名稱（如 `gpt-5`、`gpt-4o`）。程式會依模型自動選用正確的 API 參數 |
| `AI_MAX_TOKENS` | 回傳 token 上限。GPT-5 系列會先用 token 做內部推理，須留足額度（預設 8000）避免答案被截斷 |
| `DIAG_CIK` | 對指定 CIK 中「定位失敗」的 filing 印出深度診斷；不需要設 `None` |
| `SLEEP` | 對 SEC EDGAR 的禮貌間隔（秒） |
| `USER_AGENT` | SEC 要求帶可聯絡的 User-Agent，請改成自己的 |

### 建議流程
1. 先以 `LIMIT=20`、`AI_TEST_MODE=True` 跑 1 筆，確認 AI 回傳格式正確。
2. 改 `AI_TEST_MODE=False`，仍 `LIMIT=20`，跑完整測試樣本，檢查 CSV。
3. 確認無誤後改 `LIMIT=None`、`FRESH=False`，正式跑完整清單。

---

## 輸出

### `step1_sections.jsonl`
每行一筆 filing。重要欄位：`table_text`（章節全文：持股表 + 附註，餵 AI 的原料）、`table_html`（範圍內 HTML table）、`has_pledge`、`pledge_loc`（`in_section` / `out_of_section` / `none`）、`anchor_type`。

### `step3_pledges.csv`（核心輸出）
逐質押內部人一列：

| 欄位 | 說明 |
|------|------|
| `CIK` / `FILEDATE` | 來源（關聯用） |
| `insider_name` | 質押內部人姓名 |
| `share_class` | 類股別（單類股為 `Common`，多類股分開列） |
| `pledged_shares` | 質押股數 |
| `shares_owned` | 該人持股數 |
| `title` | 職位（AI 抽取；該 filing 持股表未標註職稱時可能為空） |

### `step2_owners.csv` / `step3_aggregate.csv`
彙總內部人持股（每類股一列）：`CIK`、`FILEDATE`、`share_class`、`aggregate_insider_shares`。`step2_owners.csv` 為無質押公司、`step3_aggregate.csv` 為有質押公司。

---

## AI 抽取設計（步驟 2/3）

**只送 step1 截取的章節文字（受益所有權表 + 附註），不送整份 filing。** 一份 DEF 14A 全文約數萬 token，只送相關章節（約 2,000–10,000 字）可把每筆成本壓到約一美分。

- 送給 AI 的內容 = 教授的 prompt（原文）＋ 要求嚴格回傳 JSON 的 schema 指示 ＋ step1 章節文字。
- 回傳統一 JSON schema：`aggregate_insider_ownership`（每類股一筆）與 `pledges`（每位質押內部人一筆，含姓名/類股/質押股數/持股數/職位）。
- `has_multiple_classes` 為 true 時，持股與質押皆**分類股別**報。
- 溫度設 0（舊模型）以求可重現；GPT-5 系列固定使用預設值。
- 模型相容：程式自動判斷模型家族，新模型（gpt-5、o1/o3/o4）用 `max_completion_tokens` 且不送 temperature，舊模型（gpt-4o 等）用 `max_tokens` + temperature。

職位（步驟 4）：受益所有權表若標註職稱，AI 在步驟 3 同一次呼叫即一併回傳（不增加全文 token）；若該表未標職稱，AI 回空，再由爬蟲從全文比對姓名鄰近的職稱字樣備援。部分早年 filing 的持股表本身未標職稱，職位欄因此可能為空——此為原始文件限制，非程式缺陷。

---

## 步驟 1 的定位策略

DEF 14A 受益所有權章節的標題寫法與 HTML 結構在不同年代、不同公司差異極大。最終採用：

1. **文字流截取**：把全文轉純文字，從受益所有權標題截到下一個大章節標題為止。只要質押附註落在此範圍內就一定被涵蓋，不受 HTML DOM 結構影響。
2. **評分選錨點**：同一標題字串常在目錄（TOC）、附註、文件尾段重複出現。對每個候選試截並評分，優先選含持股表表頭字樣（"Beneficially Owned"、"Amount and Nature of Beneficial"…）且有多筆股數列者，避開目錄條目（截出為空）與文件尾段雜燴（長但無表頭）。
3. **保留範圍內 HTML table**：以文字指紋比對保留，供需要欄位結構時使用。

### 已解決的三類結構問題
- **表格被拆散**：部分公司（如 Franklin Street Properties）受益所有權表被 HTML 拆成十幾張相鄰小 `<table>`。
- **純文字排版**：早年 `.txt` filing 用空格對齊排版，持股表不在 `<table>` 裡。
- **錨點誤判**：目錄條目、含標題字串的長附註、文件尾段雜燴造成錨點選錯。

---

## 程式結構

```
步驟 0   step0_load_list()              讀清單
步驟 1   step1_extract_section()        主截取
           ├ extract_text_section()         文字流截取 + 評分選錨點
           ├ collect_tables_in_text_range() 保留範圍內 HTML table
           └ is_toc_line() 等              TOC / 標題 / 附註判斷
AI       call_openai() / _parse_json()  呼叫 OpenAI、解析 JSON
步驟 2   step2_owners_no_pledge()       無質押 → 彙總持股
步驟 3   step3_pledges()                有質押 → 彙總持股 + 質押明細（含職位）
步驟 4   step4_lookup_title()           職位爬蟲備援（AI 未提供時）
診斷     diagnose()                     對定位失敗的 filing 印出結構分析
進度     load_progress / save_progress  斷點續跑
```

---

## 驗證結果（20 筆測試樣本）

涵蓋三家公司、跨多年度：
- **單類股**（CIK 1031233）：質押人、質押股數、持股數、職位皆正確。
- **多類股**（CIK 1031308 / Bentley，Class A + Class B）：持股與質押正確分類股別，單次呼叫抓出多位質押內部人。
- **多質押人**（CIK 1031316 / FSP）：McGillicuddy、MacPhee、Gribbell、Silverstein 等多人質押明細皆正確抽出。

質押三要素（姓名 / 質押股數 / 持股數）完整且正確；職位欄在少數未標註職稱的早年 filing 為空。

---

## 成本與注意事項

- 只送章節（非全文）時，每筆約一美分；完整清單（數千筆）粗估數十美元。實際以 OpenAI 官方定價為準。
- 正式跑大量資料前，確認 OpenAI 帳戶額度，並將 `FRESH` 設為 `False` 以支援中途斷線續跑。
- SEC EDGAR 對請求頻率有禮貌性要求；保留 `SLEEP` 間隔並使用可聯絡的 `USER_AGENT`。
- 沙盒或受限網路可能無法存取 `sec.gov` 與 `api.openai.com`；需在可連外環境執行。