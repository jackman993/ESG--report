# 實際 Log 內容清單

## 📋 目錄
- [Log 檔案位置](#log-檔案位置)
- [Log 1: Step 1 - 子步驟1 (碳排放計算)](#log-1-step-1---子步驟1-碳排放計算)
- [Log 2: Step 1 - 子步驟2 (TCFD 表格生成)](#log-2-step-1---子步驟2-tcfd-表格生成)
- [Log 3: Step 1 - 子步驟3 (環境章節生成)](#log-3-step-1---子步驟3-環境章節生成)
- [Log 4: Step 2 (重大議題與公司段報告)](#log-4-step-2-重大議題與公司段報告)
- [Log 5: Step 3 (治理與社會段報告)](#log-5-step-3-治理與社會段報告)

---

## 📁 Log 檔案位置

**路徑**: `C:\Users\User\Desktop\ESG_Output\_Backend\user_logs\`  
**檔案格式**: `session_{YYYYMMDD_HHMMSS}.json`  
**範例檔案**: 
- `session_20251217_193630.json` (Step 3)
- `session_20251217_193240.json` (Step 2)
- `session_20251217_192721.json` (Step 1 - 子步驟3)
- `session_20251217_192428.json` (Step 1 - 子步驟2)
- `session_20251217_192314.json` (Step 1 - 子步驟1)

---

## Log 1: Step 1 - 子步驟1 (碳排放計算)

**檔案**: `session_20251217_192314.json`  
**時間**: 2025-12-17 19:23:14

### 完整 JSON 內容：
```json
{
  "step": "Step 1 - 子步驟1",
  "industry": "焊接業",
  "monthly_bill": 90000.0,
  "company_profile": {
    "monthly_bill_ntd": 90000.0,
    "annual_revenue_ntd": 32400000.0,
    "annual_revenue_wan": 3240.0,
    "revenue_display": "3240萬元 (32,400,000)",
    "revenue_for_prompt": "3240萬元",
    "size": "中小型",
    "budget_ntd": 648000.0,
    "budget_wan": 64.8,
    "budget_display": "64.8萬元",
    "budget_for_prompt": "64.8萬元"
  },
  "emission_data": {
    "scope1": 11.63,
    "scope2": 116.35,
    "total": 127.98,
    "gasoline": 10.47,
    "refrigerant": 1.16,
    "electricity": 116.35,
    "占比": {
      "電力": 90.9,
      "車輛": 8.2,
      "冷媒": 0.9
    }
  },
  "calc_mode": "Quick (80%)",
  "session_id": "20251217_192314",
  "timestamp": "2025-12-17T19:23:14.480806"
}
```

### 欄位說明：
- **step**: 步驟識別碼
- **industry**: 產業名稱
- **monthly_bill**: 月電費（NTD）
- **company_profile**: 公司規模資訊
  - `monthly_bill_ntd`: 月電費（新台幣）
  - `annual_revenue_ntd`: 年營收（新台幣）
  - `annual_revenue_wan`: 年營收（萬元）
  - `revenue_display`: 營收顯示格式
  - `revenue_for_prompt`: 用於 Prompt 的營收格式
  - `size`: 企業規模（中小型/中型/中大型）
  - `budget_ntd`: 節能投資預算（新台幣）
  - `budget_wan`: 節能投資預算（萬元）
  - `budget_display`: 預算顯示格式
  - `budget_for_prompt`: 用於 Prompt 的預算格式
- **emission_data**: 碳排放數據
  - `scope1`: 範疇一排放量（tCO₂e）
  - `scope2`: 範疇二排放量（tCO₂e）
  - `total`: 總排放量（tCO₂e）
  - `gasoline`: 汽油排放量
  - `refrigerant`: 冷媒排放量
  - `electricity`: 電力排放量
  - `占比`: 各項排放占比（%）
- **calc_mode**: 計算模式（Quick/Detail）
- **session_id**: Session 識別碼
- **timestamp**: 時間戳記

---

## Log 2: Step 1 - 子步驟2 (TCFD 表格生成)

**檔案**: `session_20251217_192428.json`  
**時間**: 2025-12-17 19:24:28

### 完整 JSON 內容：
```json
{
  "step": "Step 1 - 子步驟2",
  "industry": "焊接業",
  "company_profile": {
    "monthly_bill_ntd": 90000.0,
    "annual_revenue_ntd": 32400000.0,
    "annual_revenue_wan": 3240.0,
    "revenue_display": "3240萬元 (32,400,000)",
    "revenue_for_prompt": "3240萬元",
    "size": "中小型",
    "budget_ntd": 648000.0,
    "budget_wan": 64.8,
    "budget_display": "64.8萬元",
    "budget_for_prompt": "64.8萬元"
  },
  "tcfd_summary": {
    "transformation_policy": "政策與法規風險",
    "transformation_raw": "政策與法規風險|||碳費徵收預計2026年實施，以每公噸500-1000元計算，焊接業年碳排約200-400公噸，年增成本20-40萬元，占營收0.6-1.2%;能源效率法規要求年節能1%以上，違規罰鍰可達50萬元;環評審查趨嚴，新廠投資時程延長6-12個月，增加前置成本15-30萬元|||導入ISO 50001能源管理系統，預算15萬元，預期年節能8-12%;更換高效率焊接設備，投資35萬元，降低用電20%;建立碳盤查機制並設定減碳目標，委外費用8萬元，符合法規要求並掌握碳足跡數據|||低碳焊接技術需求增加，傳統電弧焊接將被雷射焊接等新技術取代，設備投資需200-500萬元;客戶要求綠色供應鏈認證，取得相關認證費用10-20萬元;市場對低碳焊材需求上升，傳統焊材毛利率將下降5-10個百分點|||投資節能焊接設備64.8萬元，優先汰換老舊設備;與設備商合作開發綠色焊接解決方案，年研發預算占營收1.5%;建立綠色產品認證體系，預算12萬元，提升產品競爭力與附加價值",
    "market_trend": "消費者偏好變化風險：2026年後客戶更偏好環保焊接服務，傳統高碳排焊材需求將下降30-40%，綠色焊接技術需求年增15-20%；客戶要求ESG合規供應商，未符合環保標準將面臨訂單流失風險。",
    "market_raw": "消費者偏好變化風險：2026年後客戶更偏好環保焊接服務，傳統高碳排焊材需求將下降30-40%，綠色焊接技術需求年增15-20%；客戶要求ESG合規供應商，未符合環保標準將面臨訂單流失風險。|||營收影響：未轉型將面臨營收下滑20-25%約648-810萬元；綠色焊材成本較傳統焊材高出15-25%，短期毛利率下降3-5個百分點；客戶流失可能導致2027年後年營收減少10-15%。|||技術升級：投資64.8萬元導入低碳焊接技術;供應鏈轉型：與綠色焊材供應商建立合作關係，預計2026年完成;市場定位：發展ESG友善焊接服務，目標2027年綠色服務佔比達40%\n\n市場需求變化風險：製造業客戶受碳稅影響，2027年後對低碳焊接需求激增，傳統焊接服務市場萎縮35-45%；新能源、電動車產業焊接需求年增25-30%，但技術門檻較高，需要精密焊接能力。|||市場機會損失：未掌握新興市場將錯失年增30%業務機會約972萬元；技術落後將導致高價值訂單流失，影響營收成長15-20%;新市場進入成本較高，初期投資回收期延長至2-3年。|||市場布局：鎖定新能源產業客戶，預計2026年新市場營收佔比達25%;技術認證：取得高階焊接認證，投資預算64.8萬元用於設備升級;策略聯盟：與新能源設備商建立合作，目標2027年策略客戶佔營收35%"
  },
  "tcfd_output_folder": "C:\\Users\\User\\Desktop\\ESG_Output\\A_TCFD",
  "results_count": 5,
  "session_id": "20251217_192428",
  "timestamp": "2025-12-17T19:24:28.084327"
}
```

### 欄位說明：
- **step**: 步驟識別碼
- **industry**: 產業名稱
- **company_profile**: 公司規模資訊（同 Log 1）
- **tcfd_summary**: TCFD 摘要資訊
  - `transformation_policy`: 轉型風險政策摘要（簡短版）
  - `transformation_raw`: 轉型風險完整 LLM 回應（包含 ||| 分隔符）
  - `market_trend`: 市場風險趨勢摘要（簡短版）
  - `market_raw`: 市場風險完整 LLM 回應（包含 ||| 分隔符）
- **tcfd_output_folder**: TCFD 輸出資料夾路徑
- **results_count**: 生成的表格數量（5個）
- **session_id**: Session 識別碼
- **timestamp**: 時間戳記

### TCFD Raw 資料格式說明：
- 使用 `|||` 分隔三欄：風險描述、財務影響、因應措施
- 每欄內使用分號 `;` 分隔多個要點
- 多行資料使用換行符 `\n` 分隔

---

## Log 3: Step 1 - 子步驟3 (環境章節生成)

**檔案**: `session_20251217_192721.json`  
**時間**: 2025-12-17 19:27:21

### 完整 JSON 內容：
```json
{
  "step": "Step 1 - 子步驟3",
  "industry": "焊接業",
  "company_profile": {
    "monthly_bill_ntd": 90000.0,
    "annual_revenue_ntd": 32400000.0,
    "annual_revenue_wan": 3240.0,
    "revenue_display": "3240萬元 (32,400,000)",
    "revenue_for_prompt": "3240萬元",
    "size": "中小型",
    "budget_ntd": 648000.0,
    "budget_wan": 64.8,
    "budget_display": "64.8萬元",
    "budget_for_prompt": "64.8萬元"
  },
  "emission_data": {
    "scope1": 11.63,
    "scope2": 116.35,
    "total": 127.98,
    "gasoline": 10.47,
    "refrigerant": 1.16,
    "electricity": 116.35,
    "占比": {
      "電力": 90.9,
      "車輛": 8.2,
      "冷媒": 0.9
    }
  },
  "tcfd_output_folder": "C:\\Users\\User\\Desktop\\ESG_Output\\A_TCFD",
  "output_path": "C:\\Users\\User\\Desktop\\ESG_Output\\C_Environment\\ESG環境篇_20251217_192711.pptx",
  "summary": "本公司為年營收3240萬元的中小型焊接業者，積極建構環境治理體系以因應氣候變遷挑戰。在環境治理架構方面，公司已建立ESG管理機制，致力於環保標準合規，以滿足客戶對永續供應商的要求。\n\n碳排放管理策略著重於技術轉型升級，針對傳統高碳排焊材逐步減量使用，並投入綠色焊接技術研發，以降低整體營運碳足跡。公司預期透過製程優化與設備更新，提升能源使用效率。\n\n面對TCFD識別的氣候風險，公司制定具體因應措施：針對政策法規風險，建立法規監控機制確保合規；面對市場轉型風險，積極開發環保焊接服務，預估2026年後...",
  "test_mode": false,
  "session_id": "20251217_192721",
  "timestamp": "2025-12-17T19:27:21.115633"
}
```

### 欄位說明：
- **step**: 步驟識別碼
- **industry**: 產業名稱
- **company_profile**: 公司規模資訊（同 Log 1）
- **emission_data**: 碳排放數據（同 Log 1）
- **tcfd_output_folder**: TCFD 輸出資料夾路徑
- **output_path**: 生成的 PPTX 檔案完整路徑
- **summary**: 報告摘要（約200字，由 LLM 生成）
- **test_mode**: 是否為測試模式（false = 實際調用 LLM）
- **session_id**: Session 識別碼
- **timestamp**: 時間戳記

---

## Log 4: Step 2 (重大議題與公司段報告)

**檔案**: `session_20251217_193240.json`  
**時間**: 2025-12-17 19:32:40

### 完整 JSON 內容（舊版本）：
```json
{
  "step": "Step 2",
  "company_name": " 超級鐵工仿",
  "output_path": "C:\\Users\\User\\Desktop\\ESG_Output\\D_Company\\ESG_PPT_company_20251217_193232.pptx",
  "summary": "超級鐵工廠積極推動ESG永續發展，透過系統性重大議題分析，識別出環境保護、職業安全衛生、產品品質與客戶滿意度、員工權益保障、供應鏈管理及公司治理等六大核心議題。公司制定完整永續策略框架，以「綠色製造、安全第一、品質優先」為核心理念，建立環境管理系統降低碳排放，強化職場安全文化，並持續提升產品服務品質。\n\n在利害關係人溝通方面，建立多元化溝通管道，定期與員工、客戶、供應商、投資人及社區居民進行對話，收集回饋意見並納入營運決策考量。透過員工大會、客戶滿意度調查、供應商評鑑會議等機制，確保雙向溝通效果...",
  "session_id": "20251217_193240",
  "timestamp": "2025-12-17T19:32:40.958348"
}
```

### 更新後的 JSON 內容（新版本）：
```json
{
  "step": "Step 2",
  "company_name": "公司名稱",
  "industry": "產業名稱",
  "tcfd_market_trend": "TCFD 市場風險/趨勢摘錄",
  "output_path": "C:\\Users\\User\\Desktop\\ESG_Output\\D_Company\\ESG_PPT_company_20251217_193232.pptx",
  "summary": "報告摘要（約200字）",
  "session_id": "20251217_193240",
  "timestamp": "2025-12-17T19:32:40.958348"
}
```

### 欄位說明：
- **step**: 步驟識別碼
- **company_name**: 公司名稱（用戶輸入或預設「本公司」，可能有前後空白）
- **industry**: 產業名稱（從 Step 1 的 session_state 取得）
- **tcfd_market_trend**: TCFD 市場風險/趨勢摘錄（從 Step 1 的 tcfd_summary.market_trend 取得）
- **output_path**: 生成的 PPTX 檔案完整路徑
- **summary**: 報告摘要（約200字，由 LLM 生成）
- **session_id**: Session 識別碼
- **timestamp**: 時間戳記

### 資料來源：
- `industry`: 從 `st.session_state.get("industry_selected")` 或 `st.session_state.get("industry", "")` 取得
- `tcfd_market_trend`: 從 `st.session_state.get("tcfd_summary", {}).get("market_trend", "")` 取得

---

## Log 5: Step 3 (治理與社會段報告)

**檔案**: `session_20251217_193630.json`  
**時間**: 2025-12-17 19:36:30

### 完整 JSON 內容：
```json
{
  "step": "Step 3",
  "output_path": "C:\\Users\\User\\Desktop\\ESG_Output\\F_Governance_Social\\ESG_PPT_AB_7slides_cleanup_20251217_193621.pptx",
  "summary": "本公司建立完善的公司治理架構，董事會運作透明且職能明確，設立各專業委員會強化監督機制，確保決策程序合規有效。在風險管理方面，公司建置全面性內控制度，涵蓋營運、財務及法遵風險，透過系統化識別、評估與監控機制，有效降低企業經營風險。\n\n員工權益保障為公司核心價值，提供具競爭力的薪酬福利制度，建立公平的績效評核機制，並持續投資員工教育訓練與職涯發展。工作環境重視多元包容，保障勞工權益，營造安全健康的職場文化。\n\n社會責任實踐方面，公司積極參與社區發展，透過教育支持、環境保護、弱勢關懷等多元公益活動，發...",
  "session_id": "20251217_193630",
  "timestamp": "2025-12-17T19:36:30.091133"
}
```

### 欄位說明：
- **step**: 步驟識別碼
- **output_path**: 生成的 PPTX 檔案完整路徑
- **summary**: 報告摘要（約200字，由 LLM 生成）
- **session_id**: Session 識別碼
- **timestamp**: 時間戳記

---

## 📊 Log 結構總結

### 共同欄位：
所有 Log 都包含：
- `step`: 步驟識別碼
- `session_id`: Session 識別碼（格式：YYYYMMDD_HHMMSS）
- `timestamp`: ISO 格式時間戳記

### 各步驟特有欄位：

| 步驟 | 特有欄位 |
|------|---------|
| Step 1 - 子步驟1 | `industry`, `monthly_bill`, `company_profile`, `emission_data`, `calc_mode` |
| Step 1 - 子步驟2 | `industry`, `company_profile`, `tcfd_summary`, `tcfd_output_folder`, `results_count` |
| Step 1 - 子步驟3 | `industry`, `company_profile`, `emission_data`, `tcfd_output_folder`, `output_path`, `summary`, `test_mode` |
| Step 2 | `company_name`, `industry`, `tcfd_market_trend`, `output_path`, `summary` |
| Step 3 | `output_path`, `summary` |

### 資料流向：
```
Step 1-1 (碳排放) 
  → Step 1-2 (TCFD表格) 
    → Step 1-3 (環境章節)
      → Step 2 (公司段)
        → Step 3 (治理社會段)
```

每個步驟的資料會累積到 `session_state`，後續步驟可以使用前一步驟的資料。

---

## 🔍 注意事項

1. **資料格式**：
   - 數值欄位為浮點數或整數
   - 字串欄位可能包含前後空白（如 `company_name`）
   - 路徑使用 Windows 格式（`\\`）

2. **TCFD Raw 資料**：
   - 使用 `|||` 分隔欄位
   - 使用 `;` 分隔欄位內的多個要點
   - 使用 `\n` 分隔多行

3. **Summary 欄位**：
   - 約200字
   - 可能被截斷（以 `...` 結尾）
   - 包含換行符 `\n`

4. **時間戳記**：
   - ISO 8601 格式
   - 包含微秒精度

