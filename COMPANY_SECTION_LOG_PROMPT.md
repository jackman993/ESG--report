# Company 段（Step 2）Log 與 Prompt 清單

## 📋 目錄
- [Log 位置與結構](#log-位置與結構)
- [Company 段頁面清單](#company-段頁面清單)
- [Prompt 清單](#prompt-清單)
- [Log 與 Prompt 對應關係](#log-與-prompt-對應關係)

---

## 📁 Log 位置與結構

### Log 儲存位置
- **路徑**: `C:\Users\User\Desktop\ESG_Output\_Backend\user_logs\`
- **檔案格式**: `session_{YYYYMMDD_HHMMSS}.json`
- **讀取方式**: `env_log_reader.py` 自動讀取最新的環境段 log

### Log 讀取機制
Company 段會自動讀取 **Step 1 - 子步驟3** 的 log（環境章節生成），透過 `env_log_reader.py` 載入：
- 自動尋找最新的 log 檔案
- 標準化 log 資料格式
- 提取 prompt 上下文資訊

### Log 資料來源
Company 段使用的 log 資料來自：
- **Step 1 - 子步驟3** 的 log（`session_*.json`）
- 包含：產業、公司規模、營收、TCFD 摘要、碳排放數據

---

## 📄 Company 段頁面清單

### 主要頁面檔案

#### 1. Streamlit 頁面
**檔案**: `pages/4_📋_重大議題段報告.py`

**功能**:
- 顯示前置條件檢查（Step 1 是否完成）
- 公司名稱輸入（可選）
- 調用公司段引擎生成 PPTX
- 生成報告摘要（使用 `generate_report_summary`）
- **儲存 Log**（第 163-169 行）

**Log 內容**:
```json
{
  "step": "Step 2",
  "company_name": "公司名稱",
  "industry": "產業名稱",
  "tcfd_market_trend": "TCFD 市場風險/趨勢摘錄",
  "output_path": "生成的PPTX完整路徑",
  "summary": "報告摘要（200字）",
  "session_id": "YYYYMMDD_HHMMSS",
  "timestamp": "ISO格式時間戳記"
}
```

**Log 欄位說明**:
- `step`: 步驟識別碼
- `company_name`: 公司名稱（用戶輸入或預設「本公司」）
- `industry`: 產業名稱（從 Step 1 的 session_state 取得）
- `tcfd_market_trend`: TCFD 市場風險/趨勢摘錄（從 Step 1 的 tcfd_summary 取得）
- `output_path`: 生成的 PPTX 檔案完整路徑
- `summary`: 報告摘要（約200字，由 LLM 生成）
- `session_id`: Session 識別碼
- `timestamp`: ISO 格式時間戳記

#### 2. 引擎包裝器
**檔案**: `company_engine_wrapper_zh.py`

**功能**:
- 包裝公司段引擎
- 設置 API key 和輸出目錄
- 調用 `content_pptx_company.py` 和 `full_pptx_company.py`
- **不直接儲存 Log**（由 Streamlit 頁面儲存）

#### 3. 內容生成引擎
**檔案**: `company1.1-3.6/content_pptx_company.py`

**功能**:
- 載入環境段 log（第 37 行：`self.env_log_data = self._load_environment_log()`）
- 提取 prompt 上下文（第 38 行：`self.env_context = get_prompt_context(self.env_log_data)`）
- 生成所有頁面的內容（使用多個 prompt）
- **不直接儲存 Log**（由 Streamlit 頁面儲存）

#### 4. Log 讀取器
**檔案**: `company1.1-3.6/env_log_reader.py`

**功能**:
- 讀取最新的環境段 log
- 標準化 log 資料格式
- 提取 prompt 上下文資訊
- **不儲存 Log**（僅讀取）

---

## 🤖 Prompt 清單

### 1. Streamlit 頁面 Prompt（1個）

#### Prompt 1: Step 2 報告摘要
**位置**: `shared/utils.py` (第 202-222 行)

**Prompt**:
```
請為以下ESG重大議題與公司段報告生成200字摘要：

**公司名稱：** {company_name}

本報告涵蓋：
- 重大議題分析
- 公司永續策略
- 利害關係人溝通
- 永續發展目標與績效

請生成200字摘要，重點說明：
1. 公司識別的重大永續議題
2. 永續策略與管理方針
3. 利害關係人溝通機制
4. 具體成果與未來規劃

摘要要求：精簡、專業、突出重點，約200字。**重要：請使用純文本格式，不要使用 Markdown 標題（如 #、##）或任何格式符號，直接輸出摘要文字即可。**
```

**參數**:
- `{company_name}`: 公司名稱

**使用位置**: `pages/4_📋_重大議題段報告.py` (第 133 行)

---

### 2. Company 段內容生成 Prompts（17個）

**位置**: `company1.1-3.6/content_pptx_company.py`

所有 prompt 都會整合環境段 log 資料作為上下文。

#### Prompt 2: CEO 訊息
**函數**: `generate_ceo_message()` (第 251-266 行)

**Prompt**:
```
請撰寫約 330 字（對應 220 英文單字）的 CEO 訊息，用於 ESG 永續報告書。
平衡啟發性與責任感：歡迎讀者、概述永續願景、強調近期 ESG 成就、
承認仍面臨的挑戰，並呼籲利害關係人共同參與轉型旅程。
使用「我們」和「本公司」，保持溫暖且專業的語調，避免元評論。

公司背景：{company_context}
關鍵法規挑戰：{tcfd_policy}
```

**Log 上下文**:
- `company_context`: 公司背景（產業、營收、規模）
- `tcfd_policy`: TCFD 政策與法規風險（前300字）

---

#### Prompt 3: 合作概況
**函數**: `generate_cooperation_info()` (第 268-282 行)

**Prompt**:
```
請提供約 345 字（對應 230 英文單字）描述公司的合作概況。
重要：在第一句中使用 {COMPANY_NAME} 作為公司名稱的佔位符。
例如，以「{COMPANY_NAME} 公司擁有豐富的歷史...」或「{COMPANY_NAME} 是一家多元化...」開頭。
總結背景、商業模式、地理足跡、策略夥伴關係和組織架構。
強調使命、價值觀，以及合作如何支撐長期競爭力。
使用簡潔的中文，不使用項目符號，保持高階主管語調。

公司背景：{company_context}
```

**Log 上下文**:
- `company_context`: 公司背景

---

#### Prompt 4: 財務穩定性
**函數**: `generate_cooperation_financial()` (第 284-295 行)

**Prompt**:
```
請撰寫約 345 字（對應 230 英文單字）詳細說明財務穩定性、營收成長、資本投資和法規遵循。
討論風險管理、流動性紀律、再投資優先順序和保證機制。
保持自信且透明的語調，避免元評論，以單一流暢段落呈現敘述。

公司財務背景：{company_context}
```

**Log 上下文**:
- `company_context`: 公司背景

---

#### Prompt 5: 利害關係人識別
**函數**: `generate_stakeholder_identify()` (第 297-305 行)

**Prompt**:
```
請撰寫約 375 字（對應 250 英文單字）的利害關係人識別章節，用於 ESG 報告。
識別關鍵利害關係人群體，如投資人、客戶、員工、監管機構、供應商和社區，說明每個群體的重要性。
討論他們的期望、若被忽視的潛在風險，以及如何設定優先順序。
使用簡潔的高階主管語調，不使用項目符號或標題。
```

**Log 上下文**: 無（獨立 prompt）

---

#### Prompt 6: 利害關係人分析
**函數**: `generate_stakeholder_analysis()` (第 307-318 行)

**Prompt**:
```
請撰寫約 375 字（對應 250 英文單字）分析利害關係人的需求和影響力。
涵蓋顯著性（權力、合法性、緊迫性）、重大議題、溝通節奏，以及監控參與成效的關鍵績效指標。
保持高階主管語調，使用簡潔的中文，避免項目符號。

市場趨勢背景：{tcfd_market}
```

**Log 上下文**:
- `tcfd_market`: TCFD 市場趨勢（前300字）

---

#### Prompt 7: 重大議題文本
**函數**: `generate_material_issues_text()` (第 320-334 行)

**Prompt**:
```
請撰寫約 375 字（對應 250 英文單字）總結公司的重大議題概況。
說明優先順序設定方法、與策略和風險的連結、利害關係人意見，以及預期成果。
保持分析性且易於理解的語調，不使用元評論或項目列表。

法規與政策風險：{tcfd_policy}
市場趨勢風險：{tcfd_market}
```

**Log 上下文**:
- `tcfd_policy`: TCFD 政策與法規風險（前250字）
- `tcfd_market`: TCFD 市場趨勢風險（前250字）

---

#### Prompt 8: 重大性評估摘要
**函數**: `generate_materiality_summary()` (第 336-350 行)

**Prompt**:
```
請撰寫約 375 字（對應 250 英文單字）總結重大性評估方法和主要發現。
涵蓋雙重重大性（影響與財務）、利害關係人參與、矩陣解讀，以及由此產生的行動。
使用清晰的文字，以單一連貫的敘述呈現。

法規與政策風險：{tcfd_policy}
市場趨勢風險：{tcfd_market}
```

**Log 上下文**:
- `tcfd_policy`: TCFD 政策與法規風險（前250字）
- `tcfd_market`: TCFD 市場趨勢風險（前250字）

---

#### Prompt 9: 永續策略介紹
**函數**: `generate_sustainability_strategy_intro()` (第 352-359 行)

**Prompt**:
```
請撰寫約 165 字（對應 110 英文單字）介紹永續策略與行動計畫。
說明三大策略支柱、如何與營運整合，以及附表中描述的未來方向。
使用「我們」和「本公司」，保持高階主管語調，不使用元評論。
```

**Log 上下文**: 無（獨立 prompt）

---

#### Prompt 10: ESG 核心支柱
**函數**: `generate_esg_pillars()` (第 361-373 行)

**Prompt**:
```
請撰寫約 375 字（對應 250 英文單字）說明公司的 ESG 核心支柱，涵蓋地球（Planet）、產品（Products）和人員（People）。
討論重點領域、跨價值鏈的整合、衡量實務和問責機制。
提及表格顯示每個支柱的摘要和倡議。
使用簡潔的中文，避免項目符號。

環境績效：{emission_context}
```

**Log 上下文**:
- `emission_context`: 碳排放數據（格式：`Total annual carbon emissions: X.XX tCO₂e.`）

---

#### Prompt 11: ESG 路線圖
**函數**: `generate_esg_roadmap_context()` (第 375-389 行)

**Prompt**:
```
請撰寫約 345 字（對應 230 英文單字）敘述 ESG 路線圖。
總結基礎承諾如何演進為風險整合、營運脫碳和前瞻性合規。
強調里程碑、治理負責人，以及路線圖如何指導投資和利害關係人參與。

關鍵法規挑戰：{tcfd_policy}
環境績效現況：{emission_context}
```

**Log 上下文**:
- `tcfd_policy`: TCFD 政策與法規風險（前300字）
- `emission_context`: 碳排放數據

---

#### Prompt 12: 利害關係人溝通
**函數**: `generate_stakeholder_communication()` (第 391-401 行)

**Prompt**:
```
請提供約 375 字（對應 250 英文單字）描述利害關係人溝通實務，以補充重大性雷達圖。
涵蓋參與節奏、回饋管道、揭露承諾、升級協議，以及洞察如何影響策略決策。

市場趨勢背景：{tcfd_market}
```

**Log 上下文**:
- `tcfd_market`: TCFD 市場趨勢（前300字）

---

#### Prompt 13: SDG 摘要
**函數**: `generate_sdg_summary()` (第 403-416 行)

**Prompt**:
```
請撰寫約 375 字（對應 250 英文單字）說明公司如何將永續目標與標示的 SDG 圖示對齊。
詳細說明關鍵計畫、夥伴關係、指標和治理結構，這些將 SDG 與商業價值和利害關係人期望連結。

環境績效：{emission_context}
公司背景：{company_context}
```

**Log 上下文**:
- `emission_context`: 碳排放數據
- `company_context`: 公司背景

---

#### Prompt 14: 風險管理概述
**函數**: `generate_risk_management_overview()` (第 418-428 行)

**Prompt**:
```
請撰寫約 375 字（對應 250 英文單字）概述與重大 ESG 議題相關的企業風險管理方法。
涵蓋治理結構、評估節奏、緩解規劃、控制監控、升級和保證。

關鍵氣候相關法規風險：{tcfd_policy}
```

**Log 上下文**:
- `tcfd_policy`: TCFD 政策與法規風險（前300字）

---

#### Prompt 15-17: Governance 相關（英文，未使用 Log）
**函數**:
- `generate_governance_overview()` (第 137-140 行)
- `generate_gender_equality_overview()` (第 142-145 行)
- `generate_legal_alignment_overview()` (第 147-150 行)
- `generate_legal_appliance_overview()` (第 152-155 行)
- `generate_supervisory_board_overview()` (第 157-160 行)

**說明**: 這些是 Governance 相關的 prompt，使用英文，**不整合環境段 log**。

---

#### Prompt 18-29: Social 相關（英文，未使用 Log）
**函數**:
- `generate_social_community_investment()` (第 163-168 行)
- `generate_social_health_safety()` (第 170-175 行)
- `generate_social_diversity_policies()` (第 177-182 行)
- `generate_social_diversity_kpis()` (第 184-189 行)
- `generate_social_labor_rights()` (第 191-196 行)
- `generate_social_fair_employment()` (第 198-202 行)
- `generate_social_action_plan_overview()` (第 204-208 行)
- `generate_social_showcase_intro()` (第 210-214 行)
- `generate_social_flow_explanation()` (第 216-221 行)
- `generate_social_product_responsibility()` (第 223-228 行)
- `generate_social_customer_welfare()` (第 230-235 行)
- `generate_social_innovation()` (第 237-241 行)
- `generate_social_inclusive_economy()` (第 243-248 行)

**說明**: 這些是 Social 相關的 prompt，使用英文，**不整合環境段 log**。

---

## 🔗 Log 與 Prompt 對應關係

### Log 儲存位置

| 頁面/檔案 | Log 儲存 | Log 內容 |
|----------|---------|---------|
| `pages/4_📋_重大議題段報告.py` | ✅ 是 | Step 2 完整 log（包含公司名稱、輸出路徑、摘要） |
| `company_engine_wrapper_zh.py` | ❌ 否 | 僅包裝器，不儲存 log |
| `content_pptx_company.py` | ❌ 否 | 僅讀取 log，不儲存 |
| `env_log_reader.py` | ❌ 否 | 僅讀取 log，不儲存 |

### Prompt 使用 Log 資料

| Prompt 編號 | Prompt 名稱 | 使用 Log 資料 | Log 欄位 |
|------------|-----------|--------------|---------|
| Prompt 1 | Step 2 報告摘要 | ❌ 否 | 無 |
| Prompt 2 | CEO 訊息 | ✅ 是 | `company_context`, `tcfd_policy` |
| Prompt 3 | 合作概況 | ✅ 是 | `company_context` |
| Prompt 4 | 財務穩定性 | ✅ 是 | `company_context` |
| Prompt 5 | 利害關係人識別 | ❌ 否 | 無 |
| Prompt 6 | 利害關係人分析 | ✅ 是 | `tcfd_market` |
| Prompt 7 | 重大議題文本 | ✅ 是 | `tcfd_policy`, `tcfd_market` |
| Prompt 8 | 重大性評估摘要 | ✅ 是 | `tcfd_policy`, `tcfd_market` |
| Prompt 9 | 永續策略介紹 | ❌ 否 | 無 |
| Prompt 10 | ESG 核心支柱 | ✅ 是 | `emission_context` |
| Prompt 11 | ESG 路線圖 | ✅ 是 | `tcfd_policy`, `emission_context` |
| Prompt 12 | 利害關係人溝通 | ✅ 是 | `tcfd_market` |
| Prompt 13 | SDG 摘要 | ✅ 是 | `emission_context`, `company_context` |
| Prompt 14 | 風險管理概述 | ✅ 是 | `tcfd_policy` |
| Prompt 15-17 | Governance 相關 | ❌ 否 | 無 |
| Prompt 18-29 | Social 相關 | ❌ 否 | 無 |

### Log 資料流向

```
Step 1 - 子步驟3 Log
  ↓
env_log_reader.py (讀取最新 log)
  ↓
content_pptx_company.py (載入 log 資料)
  ↓
get_prompt_context() (提取上下文)
  ↓
各 Prompt 函數 (使用上下文生成內容)
  ↓
生成 PPTX
  ↓
pages/4_📋_重大議題段報告.py (儲存 Step 2 Log)
```

---

## 📊 統計資訊

- **總 Log 儲存位置**: 1 個（Streamlit 頁面）
- **總 Prompt 數量**: 29 個
  - Streamlit 頁面: 1 個
  - Company 段內容生成: 28 個
    - 使用 Log 資料: 13 個
    - 不使用 Log 資料: 15 個（Governance + Social）
- **Log 讀取位置**: 1 個（`env_log_reader.py`）
- **Log 資料來源**: Step 1 - 子步驟3 的環境段 log

---

## 📝 備註

1. **Log 讀取機制**:
   - Company 段會自動讀取最新的環境段 log
   - 如果找不到 log，會使用空值或預設值

2. **Prompt 上下文**:
   - `company_context`: 公司背景（產業、營收、規模）
   - `tcfd_policy`: TCFD 政策與法規風險
   - `tcfd_market`: TCFD 市場趨勢
   - `emission_context`: 碳排放數據

3. **語言**:
   - Company 段主要頁面（Prompt 2-14）使用繁體中文
   - Governance 和 Social 相關（Prompt 15-29）使用英文

4. **字數要求**:
   - 中文約 1.5 字 = 1 英文單字
   - 例如：220 英文單字 ≈ 330 中文字

