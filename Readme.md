# SEC DEF 14A 質押分析 Pipeline

從 SEC EDGAR 的 DEF 14A（委託書 / proxy statement）中，定位並截取「公司管理階層、主要股東及董事持股」（受益所有權，Item 403）章節，判斷公司是否有內部人質押，並透過 AI（OpenAI API）抽取出彙總內部人持股與逐人質押明細。每筆並附「信心標記」，低信心者自動列入待複查清單。

5 個步驟全部接通，並在 20 筆測試樣本上驗證（含多類股、多質押人、機構表 vs 內部人表等複雜情況）。

---

## 流程概觀

| 步驟 | 內容 | 狀態 |
|------|------|------|
| 步驟 0 | 依清單掃描 DEF 14A filings（以文件 URL 去重），抓取原始檔案 | ✅ |
| 步驟 1 | 定位受益所有權章節，截取「表格 + 附註」，判斷是否含質押，並評估信心 | ✅ |
| 步驟 2 | 【無質押公司】把章節餵 AI → 彙總內部人持股（多類股分開） | ✅ |
| 步驟 3 | 【有質押公司】把章節餵 AI → 彙總持股 + 逐人質押明細（姓名/質押股數/持股數/職位，多類股分開） | ✅ |
| 步驟 4 | 質押內部人職位：步驟 3 由 AI 一併回傳；AI 未提供時以爬蟲從全文備援 | ✅ |

> 步驟 2/3 的 prompt 由指導教授提供，程式原文照用。步驟 3 是步驟 2 的超集。

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

`.gitignore` 已排除 `.env`。另附 `.env.example` 作為範本（可安全推 git）。程式啟動時以 `python-dotenv` 自動載入。

---

## 輸入

清單檔（預設 `DEF14A2023_PLEDGE_RA2_202605.xlsx`）。程式只用到 `CIK`、`FILEDATE`、`http` 三欄（大小寫不拘、支援別名）。

> **重要：以 `http`（文件 URL）去重，不以 CIK+FILEDATE。** 清單中同一份 filing 會因「多個內部人各佔一列」而重複出現（同 CIK+FILEDATE），但那是同一份文件，只需抓一次、AI 一次回傳所有質押人。若用 CIK+FILEDATE 去重，會誤刪「同公司同日但不同文件」的情況（實測 3063 vs 3064，差 1 筆）。完整清單去重後約 3,064 份不重複 filing。
>
> 此清單為**質押研究樣本**，前段有部分人工登記的答案欄（FULLNAME / PLEDGE_SHARES / SHROWN_TOT 等），後段多為空白。本專案的目標是用 AI 自動重抽全部資料（不同於先前以人工為主的專案）。

---

## 執行

```
python "Def14a pipeline.py"
```

主要設定（檔案開頭）：

| 設定 | 用途 |
|------|------|
| `LIMIT` | 試跑筆數；正式跑改 `None` |
| `FRESH` | `True` 忽略舊進度檔整批重跑；資料量大時改 `False` 以支援斷點續跑 |
| `RUN_AI` | `False` 時只做步驟 0/1（不呼叫 AI） |
| `AI_TEST_MODE` | `True` 只跑 1 筆並印出「送什麼 / 收什麼」，驗證 prompt 與 schema |
| `OPENAI_MODEL` | 模型名稱（如 `gpt-5`、`gpt-4o`）。程式依模型自動選用正確 API 參數 |
| `AI_MAX_TOKENS` | 回傳 token 上限（預設 8000）。GPT-5 系列會先用 token 做內部推理，須留足額度避免截斷 |
| `DIAG_CIK` / `SLEEP` / `USER_AGENT` | 診斷對象 CIK / EDGAR 禮貌間隔 / 你的可聯絡 User-Agent |

### 建議流程
1. `LIMIT=20`、`AI_TEST_MODE=True` 跑 1 筆，確認 AI 回傳格式。
2. `AI_TEST_MODE=False`、`LIMIT=20` 跑測試樣本，看信心統計與 CSV。
3. 中等規模預跑（如 `LIMIT=200`）估算高信心比例與成本。
4. 確認後 `LIMIT=None`、`FRESH=False` 跑完整清單。

---

## 輸出

### `step1_sections.jsonl`
每行一筆 filing。重要欄位：`table_text`（章節全文，餵 AI 的原料）、`table_html`、`has_pledge`、`pledge_loc`（`in_section` / `out_of_section` / `none`）、`anchor_type`、`confidence`（high/medium/low）、`review_reason`。

### `step3_pledges.csv`（核心輸出）
逐質押內部人一列：`CIK`、`FILEDATE`、`insider_name`、`share_class`（單類股為 `Common`，多類股分列）、`pledged_shares`、`shares_owned`、`title`（職位，AI 抽取；持股表未標職稱時可能為空）。

### `step2_owners.csv` / `step3_aggregate.csv`
彙總內部人持股（每類股一列）：`CIK`、`FILEDATE`、`share_class`、`aggregate_insider_shares`。前者為無質押公司、後者為有質押公司。

### `needs_review.csv`（待複查清單）
所有 medium / low 信心、以及抓取失敗的 filing，附 `confidence` 與 `reason`（如「疑似只抓到機構大股東表」「質押在章節外，AI 原料可能漏質押」「抓取失敗（需重試）」）。供人工複查。

---

## 信心標記機制

完整清單橫跨二十餘年、上百家公司，filing 排版變體屬長尾分布，無法在執行前窮舉修正所有個案。因此程式對每筆自評信心，讓多數高信心者自動完成、少數低信心者標記待人工複查，而非靜默產出可能錯誤的資料。

判斷依據（程式可自行判定）：
- 章節是否成功定位、長度是否正常（過短/過長皆可疑）。
- 抓到的是「內部人持股表」還是「只有 5% 機構大股東表」（質押人必為內部人，僅抓到機構表 → 低信心）。
- `has_pledge=True` 但質押不在截取的章節內（餵 AI 的原料可能漏質押 → 低信心）。

等級：`high`（直接採用）、`medium`（輕微異常，如長度偏短）、`low`（疑似抓錯表或漏質押，務必人工複查）。執行結束會印 high/medium/low 統計。

> 20 筆測試樣本結果：high 16、medium 0、low 3、抓取失敗 1。低信心者正確指向特定公司（1032033）較刁鑽年度的 filing。

---

## AI 抽取設計（步驟 2/3）

**只送 step1 截取的章節文字（受益所有權表 + 附註），不送整份 filing**，把每筆成本壓到約一美分。送給 AI 的內容 = 教授的 prompt（原文）＋ JSON schema 指示 ＋ 章節文字。

- 回傳統一 JSON schema：`aggregate_insider_ownership`（每類股一筆）與 `pledges`（每位質押內部人：姓名/職位/類股/質押股數/持股數）。
- `has_multiple_classes` 為 true 時，持股與質押皆**分類股別**報。
- 模型相容：程式自動判斷模型家族——新模型（gpt-5、o1/o3/o4）用 `max_completion_tokens` 且不送 temperature；舊模型（gpt-4o 等）用 `max_tokens` + temperature（設 0 求可重現）。
- 職位：持股表若標職稱，AI 於步驟 3 同一次呼叫一併回傳；未標則由爬蟲從全文備援。部分早年 filing 持股表未標職稱，職位欄可能為空——此為原始文件限制。

---

## 步驟 1 的定位策略

受益所有權章節的標題寫法與 HTML 結構在不同年代、公司差異極大。採用：

1. **文字流截取**：全文轉純文字，從受益所有權標題截到下一個大章節標題。質押附註只要落在此範圍內就被涵蓋，不受 DOM 結構影響。
2. **以「標題」為主的評分選錨點**：同一標題字串常在目錄、附註、文件尾段重複出現，且揭露常分「內部人表」與「5% 機構大股東表」兩段。評分主依據是標題文字——明指「by directors and executive officers」等內部人表者強加分；「ownership guidelines」「how to vote」等雜訊標題強扣分；截出內容含質押字樣再加分。比從截出內容猜測可靠（截出範圍可能過大而使內容特徵失真）。
3. **相鄰行合併比對標題**：標題可能被換行拆成多行（如 "Principal" / "Shareholders"），故同時比對單行與相鄰 2–3 行合併。
4. **保留範圍內 HTML table**：以文字指紋比對保留，供需要欄位結構時使用。

### 已解決的結構問題（逐一在擴大樣本時發現並修正）
- 表格被拆成十幾張相鄰小 `<table>`（Franklin Street Properties）。
- 早年 `.txt` filing 用空格對齊排版，持股表不在 `<table>` 裡。
- 目錄條目、含標題字串的長附註、文件尾段雜燴造成錨點誤判。
- 標題被換行拆開（"Principal" / "Shareholders"）。
- 「5% 機構大股東表」與「內部人持股表」分開，誤抓到機構表（質押不在其中）。

> 仍可能有未見過的排版變體；這正是信心標記機制存在的原因——未能可靠處理者會被標為低信心、列入待複查，而非靜默出錯。

---

## 程式結構

```
步驟 0   step0_load_list()              讀清單（以 http 去重）
步驟 1   step1_extract_section()        主截取
           ├ extract_text_section()         文字流截取 + 標題評分選錨點
           ├ collect_tables_in_text_range() 保留範圍內 HTML table
           └ _assess_confidence()           信心評估
AI       call_openai() / _parse_json()  呼叫 OpenAI、解析 JSON
步驟 2   step2_owners_no_pledge()       無質押 → 彙總持股
步驟 3   step3_pledges()                有質押 → 彙總持股 + 質押明細（含職位）
步驟 4   step4_lookup_title()           職位爬蟲備援（AI 未提供時）
診斷     diagnose()                     對定位失敗的 filing 印出結構分析
進度     load_progress / save_progress  以 http 為鍵的斷點續跑
```

---

## 成本與注意事項

- 只送章節（非全文）時，每筆約一美分；完整清單（約 3,064 份）粗估數十美元。實際以 OpenAI 官方定價為準。
- 目前低信心的 filing 仍會呼叫 AI（其原料可能不準）；如需省費，可改為「低信心先跳過 AI」（尚未啟用）。
- 正式跑大量資料前確認 OpenAI 額度，並將 `FRESH` 設為 `False` 以支援中途斷線續跑。
- SEC EDGAR 對請求頻率有禮貌性要求；保留 `SLEEP` 間隔並使用可聯絡的 `USER_AGENT`。
- 沙盒或受限網路可能無法存取 `sec.gov` 與 `api.openai.com`；需在可連外環境執行。