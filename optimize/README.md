# 用於優化推薦方法

合併計畫名稱、關鍵字與摘要，優化成「三者分別比對相似度再加總」的做法。
摘要會利用 local model（GPT-oss-120B）先拆解成四個欄位（應用方向 / 欲解決問題 / 達成目標 / 解決方法）再比對。

> 這裡是先前優化測試使用的程式碼，所有需要的資料均存於 `data/`。

## 檔案說明

| 檔案 | 說明 |
| --- | --- |
| `abstract.py` | 摘要拆解（用於過往通過案件與當次申請案件）|
| `store_vectordb.py` / `store_vectordb_by_project.py` | 將過往通過案件存到向量資料庫。前者以「教授」為單位，後者以「專案」為單位（依名稱、關鍵字與摘要存成兩個 vectordb、7 個 collection）|
| `apply_result.py` / `apply_result_by_project.py` | 比對申請案件與通過案件計算相似度分數。前者以「教授」為單位，後者以「專案」為單位 |
| `recommand.py` / `recommand_by_project.py` | 依比對結果輸出推薦委員（去掉共同研究人員、同校、指導教授等利益迴避對象）。前者以「教授」為單位，後者以「專案」為單位 |
| `save_commitee_data.py` | 取得過往通過與當次申請案件的教授名稱，供後續爬蟲使用 |
| `crawler_advisor.py` | 於 NDLTD（碩博士論文加值網）取得教授的指導教授 |
| `crawler_degree.py` | 於 NSTC（國家科學及技術委員會研究人才查詢）以「教授名稱＋當前任職單位」取得教授畢業院校（註：部分教授未公開，故查不到）|

## 執行方法

先啟用環境：

```sh
conda activate MOST
```

下列各步驟的檔案路徑、sheet 名稱、資料庫路徑等都可以用命令列參數覆寫。
**不加任何參數就會使用程式內建的預設值**；要換一批資料時，只需在指令後面加上對應參數即可，不用改程式碼。
每支程式都支援 `-h` / `--help` 查看完整參數說明，例如：

```sh
python apply_result_by_project.py -h
```

### 1. 儲存向量資料庫

```sh
# 先把過往通過案件的摘要拆解成四欄位
python abstract.py \
  --input  data/research_proj/115計算機學門審查/pass_project.xlsx \
  --sheets 108 109 110 111 112 113 114 \
  --output data/research_proj/115計算機學門審查/pass_project_with_abstract.xlsx

# 再把拆解結果存入向量資料庫（以專案為單位）
# --type research / industry 會自動決定資料庫路徑，不用自己打
python store_vectordb_by_project.py \
  --type   research \
  --input  data/research_proj/115計算機學門審查/pass_project_with_abstract.xlsx \
  --sheets 108 109 110 111 112 113 114
```

### 2. 爬取教授資料

```sh
python save_commitee_data.py
python crawler_degree.py
python crawler_advisor.py
```

### 3. 輸出推薦委員

```sh
# 先拆解「當次申請案件」的摘要
python abstract.py \
  --input  data/research_proj/115計算機學門審查/新興大專生計畫/115新興大專生計畫.xlsx \
  --sheets 工作表1 \
  --output data/research_proj/115計算機學門審查/新興大專生計畫/apply_with_abstract.xlsx

# 比對申請案件與資料庫、計算相似度分數（會自動比對全部欄位）
# --type 需與上面 store 步驟相同，才會讀到正確的資料庫
python apply_result_by_project.py \
  --type   research \
  --input  data/research_proj/115計算機學門審查/新興大專生計畫/apply_with_abstract.xlsx \
  --sheets 工作表1 \
  --output data/research_proj/115計算機學門審查/新興大專生計畫/result_score.xlsx

# 依分數輸出推薦委員並做利益迴避過濾
python recommand_by_project.py \
  --type          research \
  --apply         data/research_proj/115計算機學門審查/新興大專生計畫/115新興大專生計畫.xlsx \
  --sheet         工作表1 \
  --score         data/research_proj/115計算機學門審查/新興大專生計畫/result_score.xlsx \
  --committee     data/RDF_database/committee_all_education_with_advisor.xlsx \
  --blacklist     data/retiree_blacklist.csv \
  --org-output    data/research_proj/115計算機學門審查/新興大專生計畫/recommendation_results_org_colored.xlsx \
  --filter-output data/research_proj/115計算機學門審查/新興大專生計畫/recommendation_results_filter.xlsx
```

## 換一批資料時，需要自行更改的地方

以下整理每支程式「使用者需要自行提供／確認」的參數（皆有預設值，可只覆寫需要的項目）。

### `abstract.py`（摘要拆解）

| 參數 | 說明 |
| --- | --- |
| `--input` | 要拆解摘要的 Excel 檔（過往通過案件或當次申請案件）|
| `--sheets` | 要處理的工作表名稱，可一次給多個（例：`--sheets 108 109 110`）|
| `--output` | 拆解結果輸出的 Excel 檔 |
| `--ollama-host` | Ollama 服務位址（預設 `http://localhost:1228`）|
| `--model` | local LLM 模型名稱（預設 `gpt-oss:120b`）|

### `store_vectordb_by_project.py`（建立向量資料庫）

| 參數 | 說明 |
| --- | --- |
| `--input` | 過往通過案件（已含摘要拆解結果）的 Excel 檔 |
| `--sheets` | 要讀取的工作表名稱，通常是年度 |
| `--type` | 計畫類型 `research` / `industry`，自動決定資料庫路徑（預設 `research`）|
| `--basic-db` / `--abstract-db` | 進階：手動指定資料庫路徑，覆寫 `--type` 的預設值（一般不需要）|

### `apply_result_by_project.py`（計算相似度）

> 一律比對全部欄位（名稱、關鍵字、摘要拆解出的四個欄位），不需也不能只選部分欄位。

| 參數 | 說明 |
| --- | --- |
| `--input` | 當次申請案件的 Excel 檔 |
| `--sheets` | 申請案件要讀取的工作表名稱 |
| `--type` | 計畫類型 `research` / `industry`，需與 store 步驟相同（預設 `research`）|
| `--basic-db` / `--abstract-db` | 進階：手動指定資料庫路徑，覆寫 `--type` 的預設值（一般不需要）|
| `--output` | 相似度分數輸出的 Excel 檔 |

### `recommand_by_project.py`（輸出推薦委員）

> 一律納入全部欄位計分，權重寫在程式的 `calculate_score()` 內。

| 參數 | 說明 |
| --- | --- |
| `--apply` | 當次申請案件的 Excel 檔（取申請人／共同主持人／機構做利益迴避）|
| `--sheet` | 申請案件要讀取的工作表名稱 |
| `--score` | `apply_result_by_project.py` 產出的相似度分數 Excel 檔 |
| `--committee` | 委員畢業學校＋博士指導教授資料（爬蟲產出）|
| `--blacklist` | 退休或黑名單委員名單 csv |
| `--type` | 計畫類型：`industry`(產學合作) 會把本次申請者也列入利益迴避；`research`(研究計畫) 則不會 |
| `--org-output` | 未過濾、含上色標記的推薦結果輸出檔 |
| `--filter-output` | 已做利益迴避過濾後的乾淨推薦結果輸出檔 |

### ⚠️ 注意事項

- **`--type` 在「建立資料庫」與「計算相似度」兩步驟要相同**（同為 `research` 或同為 `industry`），否則會讀到空的或錯誤的資料庫。
- 比對欄位固定為全部 6 個欄位，不開放挑選，`apply` 與 `recommand` 兩步驟自然一致。
- `abstract.py` 需先啟動 Ollama 服務並載入對應模型（`--ollama-host` / `--model`）。
