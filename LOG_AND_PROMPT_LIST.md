# ESG 報告系統 - Log 清單與對應 Prompt

## 📋 目錄
1. [Log 清單](#log-清單)
2. [Prompt 清單](#prompt-清單)
3. [對應關係](#對應關係)

---

## 📝 Log 清單

### Log 儲存位置
- **路徑**: `C:\Users\User\Desktop\ESG_Output\_Backend\user_logs\`
- **檔案格式**: `session_{YYYYMMDD_HHMMSS}.json`
- **檔案結構**: 包含 `session_id`, `timestamp`, 以及各步驟的資料

---

### Step 1 - 子步驟1: 碳排放計算
**檔案**: `1_🌍_碳排與TCFD氣候治理.py` (第 294-302 行)

**Log 內容**:
```json
{
  "step": "Step 1 - 子步驟1",
  "industry": "產業名稱",
  "monthly_bill": 月電費金額,
  "company_profile": {
    "monthly_bill_ntd": 月電費,
    "annual_revenue_ntd": 年營收,
    "annual_revenue_wan": 年營收(萬元),
    "revenue_display": "顯示格式",
    "revenue_for_prompt": "提示用格式",
    "size": "企業規模",
    "budget_ntd": 預算,
    "budget_wan": 預算(萬元),
    "budget_display": "預算顯示",
    "budget_for_prompt": "預算提示"
  },
  "emission_data": {
    "scope1": 範疇一排放,
    "scope2": 範疇二排放,
    "total": 總排放,
    "gasoline": 汽油排放,
    "refrigerant": 冷媒排放,
    "electricity": 電力排放,
    "占比": {...}
  },
  "calc_mode": "Quick (80%)" 或 "Detail (95%)"
}
```

**對應 Prompt**: ❌ 無（此步驟不使用 LLM）

---

### Step 1 - 子步驟2: TCFD 表格生成
**檔案**: `1_🌍_碳排與TCFD氣候治理.py` (第 461-469 行)

**Log 內容**:
```json
{
  "step": "Step 1 - 子步驟2",
  "industry": "產業名稱",
  "company_profile": {...},
  "tcfd_summary": {
    "transformation_policy": "轉型風險摘要",
    "transformation_raw": "原始LLM回應",
    "market_trend": "市場風險摘要",
    "market_raw": "原始LLM回應"
  },
  "tcfd_output_folder": "TCFD輸出資料夾路徑",
  "results_count": 5
}
```

**對應 Prompt**: ✅ 5個 TCFD 表格 Prompt（見下方）

---

### Step 1 - 子步驟3: 環境章節生成
**檔案**: `1_🌍_碳排與TCFD氣候治理.py` (第 668-678 行)

**Log 內容**:
```json
{
  "step": "Step 1 - 子步驟3",
  "industry": "產業名稱",
  "company_profile": {...},
  "emission_data": {...},
  "tcfd_output_folder": "TCFD輸出資料夾路徑",
  "output_path": "生成的PPTX完整路徑",
  "summary": "報告摘要（200字）",
  "test_mode": false
}
```

**對應 Prompt**: ✅ Step 1 摘要 Prompt（見下方）

---

### Step 2: 重大議題與公司段報告
**檔案**: `4_📋_重大議題段報告.py` (第 163-169 行)

**Log 內容**:
```json
{
  "step": "Step 2",
  "company_name": "公司名稱",
  "output_path": "生成的PPTX完整路徑",
  "summary": "報告摘要（200字）"
}
```

**對應 Prompt**: ✅ Step 2 摘要 Prompt（見下方）

---

### Step 3: 治理與社會段報告
**檔案**: `5_🏛️_治理與社會報告.py` (第 148-153 行)

**Log 內容**:
```json
{
  "step": "Step 3",
  "output_path": "生成的PPTX完整路徑",
  "summary": "報告摘要（200字）"
}
```

**對應 Prompt**: ✅ Step 3 摘要 Prompt（見下方）

---

## 🤖 Prompt 清單

### 1. TCFD 表格生成 Prompts（5個）

**位置**: `1_🌍_碳排與TCFD氣候治理.py` (第 76-147 行)

**專家角色**:
```
你是 ESG 的 GRI 和 TCFD 專家。
```

#### Prompt 1: 01 轉型風險
```
你是 ESG 的 GRI 和 TCFD 專家。針對「{industry}」進行 TCFD 轉型風險分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：政策與法規風險
第2行：綠色產品與科技風險
只輸出 2 行，不要其他文字。
```

**參數**:
- `{industry}`: 產業名稱
- `{revenue}`: 年營收（萬元格式）
- `{budget}`: 節能投資預算（萬元格式）

---

#### Prompt 2: 02 市場風險
```
你是 ESG 的 GRI 和 TCFD 專家。針對「{industry}」進行 TCFD 市場風險分析，聚焦 2026 年以後趨勢，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：消費者偏好變化風險
第2行：市場需求變化風險
只輸出 2 行，不要其他文字。
```

**參數**: 同上

---

#### Prompt 3: 03 實體風險
```
你是 ESG 的 GRI 和 TCFD 專家。針對「{industry}」進行 TCFD 實體風險分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：極端氣候事件風險
第2行：長期氣候變遷風險
只輸出 2 行，不要其他文字。
```

**參數**: 同上

---

#### Prompt 4: 04 溫升風險
```
你是 ESG 的 GRI 和 TCFD 專家。針對「{industry}」進行 TCFD 溫升情境風險分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：升溫1.5°C情境風險
第2行：升溫2°C以上情境風險
只輸出 2 行，不要其他文字。
```

**參數**: 同上

---

#### Prompt 5: 05 資源效率
```
你是 ESG 的 GRI 和 TCFD 專家。針對「{industry}」進行 TCFD 資源效率機會分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
機會描述|||潛在效益|||行動方案
第1行：能源效率提升機會
第2行：資源循環利用機會
只輸出 2 行，不要其他文字。
```

**參數**: 同上

---

### 2. 報告摘要 Prompts（3個）

**位置**: `shared/utils.py` (第 146-240 行)

#### Prompt 6: Step 1 環境段摘要
```
請為以下ESG環境段報告生成200字摘要：

**產業類別：** {industry}
**公司規模：** {size}
**年度營收：** {annual_revenue_wan}萬元

**碳排放數據：**
{total_emission}

**TCFD氣候風險摘要：**
- 轉型風險：{transformation_policy前100字}
- 市場風險：{market_trend前100字}

請生成200字摘要，重點說明：
1. 公司的環境治理架構
2. 碳排放管理策略
3. TCFD氣候風險因應措施
4. 永續發展目標與成果

摘要要求：精簡、專業、突出重點，約200字。**重要：請使用純文本格式，不要使用 Markdown 標題（如 #、##）或任何格式符號，直接輸出摘要文字即可。**
```

**參數**:
- `{industry}`: 產業類別
- `{size}`: 公司規模
- `{annual_revenue_wan}`: 年度營收（萬元）
- `{total_emission}`: 總碳排放數據
- `{transformation_policy}`: 轉型風險摘要（前100字）
- `{market_trend}`: 市場風險摘要（前100字）

---

#### Prompt 7: Step 2 公司段摘要
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

---

#### Prompt 8: Step 3 治理與社會段摘要
```
請為以下ESG治理與社會段報告生成200字摘要：

本報告涵蓋：
- 公司治理架構與運作
- 董事會職能與監督機制
- 風險管理與內控制度
- 社會責任與員工權益
- 社區參與與社會貢獻

請生成200字摘要，重點說明：
1. 公司治理架構與運作機制
2. 風險管理與內控制度
3. 員工權益與職場環境
4. 社會責任與社區參與

摘要要求：精簡、專業、突出重點，約200字。**重要：請使用純文本格式，不要使用 Markdown 標題（如 #、##）或任何格式符號，直接輸出摘要文字即可。**
```

**參數**: 無動態參數

---

### 3. TCFD 報告生成器 Prompt

**位置**: `2_🏭_TCFD報告生成器.py` (第 373 行)

#### Prompt 9: TCFD 完整報告（JSON格式）
```
請為「{industry_input}」產業生成一份 TCFD 氣候風險分析報告。

請嚴格按照以下 JSON 格式輸出：

```json
{
    "industry": "{industry_input}",
    "risks": [
        {
            "description": "風險1標題\n詳細描述...",
            "impact": "影響1標題\n詳細影響...",
            "actions": "措施1標題\n詳細措施..."
        },
        ...
    ],
    "action_plans": [
        {"name": "方案名稱1", "measure": "具體措施", "timeline": "2024-2025", "priority": "高"},
        ...
    ]
}
```

請確保輸出為有效的 JSON 格式。
```

**參數**:
- `{industry_input}`: 產業名稱

---

## 🔗 對應關係

| Log 步驟 | 使用的 Prompts | 說明 |
|---------|---------------|------|
| Step 1 - 子步驟1 | ❌ 無 | 碳排放計算，不使用 LLM |
| Step 1 - 子步驟2 | Prompt 1-5 | 生成 5 個 TCFD 表格 |
| Step 1 - 子步驟3 | Prompt 6 | 生成環境段報告摘要 |
| Step 2 | Prompt 7 | 生成公司段報告摘要 |
| Step 3 | Prompt 8 | 生成治理與社會段報告摘要 |
| TCFD 報告生成器 | Prompt 9 | 生成完整 TCFD 報告（獨立功能） |

---

## 📊 統計資訊

- **總 Log 類型**: 5 種
- **總 Prompt 數量**: 9 個
- **TCFD 表格 Prompts**: 5 個
- **報告摘要 Prompts**: 3 個
- **獨立功能 Prompts**: 1 個

---

## 📝 備註

1. **Log 檔案命名**: `session_{YYYYMMDD_HHMMSS}.json`
2. **Log 儲存位置**: `C:\Users\User\Desktop\ESG_Output\_Backend\user_logs\`
3. **Prompt 參數**: 所有 Prompt 都使用 `.format()` 方法進行參數替換
4. **模型**: 所有 API 調用都使用 `claude-sonnet-4-20250514` 模型
5. **輸出格式**: TCFD 表格使用 `|||` 分隔符，摘要使用純文本格式

