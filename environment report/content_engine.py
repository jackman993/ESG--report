"""
ESG 報告生成器 - 內容生成引擎（環境篇專用）
"""
import anthropic
import re
from config import ANTHROPIC_API_KEY, CLAUDE_MODEL


class ContentEngine:
    """使用 Claude 生成環境篇報告內容"""

    def __init__(self, test_mode=False, company_profile=None, api_key=None):
        self.test_mode = test_mode
        self.company_profile = company_profile or {}
        
        # 使用傳入的 API Key，否則用 config 的
        actual_api_key = api_key or ANTHROPIC_API_KEY
        
        if not test_mode and actual_api_key:
            try:
                self.client = anthropic.Anthropic(api_key=actual_api_key)
                print(f"  ✓ ContentEngine 已初始化（API Key: {actual_api_key[:10]}...）")
            except Exception as e:
                print(f"  ✗ ContentEngine 初始化失敗：{e}")
                self.client = None
        else:
            self.client = None
            if test_mode:
                print("  ⚠ 測試模式：跳過 API 呼叫")
            elif not actual_api_key:
                print("  ⚠ 警告：未提供 API Key")
    
    def _get_company_context(self):
        """取得公司規模背景描述"""
        if not self.company_profile:
            return ""
        
        size = self.company_profile.get("size", "中小型")
        revenue = self.company_profile.get("revenue_display", "未知")
        budget = self.company_profile.get("budget_display", "未知")
        
        return f"""
企業規模背景：
- 企業規模：{size}企業
- 推估年營收：{revenue}元
- 建議節能投資預算：{budget}元
請根據此規模，調整描述語氣和建議金額。
"""

    def _clean_llm_output(self, text):
        """清理 LLM 輸出中的元指令和標籤"""
        if not text:
            return text
        
        # 定義需要移除的開頭模式（中英文）
        patterns_to_remove = [
            r'^以下是.*?[：:]\s*',           # 以下是...：
            r'^以下為.*?[：:]\s*',           # 以下為...：
            r'^这是.*?[：:]\s*',             # 这是...：
            r'^這是.*?[：:]\s*',             # 這是...：
            r'^內容如下.*?[：:]\s*',         # 內容如下：
            r'^内容如下.*?[：:]\s*',         # 内容如下：
            r'^Here is the.*?[：:]\s*',     # Here is the...
            r'^Here\'s the.*?[：:]\s*',     # Here's the...
            r'^The following is.*?[：:]\s*', # The following is...
            r'^Below is.*?[：:]\s*',         # Below is...
            r'^以下.*?[：:]\s*',             # 以下：
            r'^\【.*?\】\s*',                # 【說明】
            r'^「.*?」\s*',                  # 「標題」
            r'^\*\*.*?\*\*\s*',              # **標題**
            r'^##\s+.*?\n',                  # ## 標題
            r'^#\s+.*?\n',                   # # 標題
        ]
        
        # 逐一套用清理模式
        cleaned_text = text
        for pattern in patterns_to_remove:
            cleaned_text = re.sub(pattern, '', cleaned_text, flags=re.IGNORECASE | re.MULTILINE)
        
        # 移除開頭和結尾的多餘空白
        cleaned_text = cleaned_text.strip()
        
        # 移除多個連續換行（保留最多一個空行）
        cleaned_text = re.sub(r'\n{3,}', '\n\n', cleaned_text)
        
        return cleaned_text

    def generate(self, prompt, max_tokens=1000):
        """呼叫 Claude API 並清理輸出"""
        # 測試模式：返回佔位文字
        if self.test_mode:
            return "[測試模式] 此處為 LLM 生成內容。正式執行時將由 Claude API 生成專業 ESG 報告內容。本公司致力於永續發展，積極落實環境保護政策。"
        
        # 檢查 client 是否初始化
        if not self.client:
            error_msg = "[內容生成失敗：API Key 未提供或客戶端未初始化]"
            print(f"✗ {error_msg}")
            return error_msg
        
        try:
            message = self.client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=max_tokens,
                messages=[{
                    "role": "user",
                    "content": prompt
                }]
            )
            # ✅ 在返回前清理輸出
            raw_text = message.content[0].text
            cleaned_text = self._clean_llm_output(raw_text)
            return cleaned_text
        except Exception as e:
            error_msg = f"[內容生成失敗：{e}]"
            print(f"✗ Claude API 錯誤：{e}")
            return error_msg

    # ==================== 環境篇內容生成方法 ====================

    def generate_environmental_cover(self, config):
        """環境篇封面引言 - 275字"""
        prompt = """為ESG報告環境篇寫約275字引言，包含：氣候變遷挑戰、企業環境責任、永續發展承諾，建立TCFD。語調專業溫暖。使用「我們」「本公司」等詞彙，避免「這間公司」「該企業」等第三人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_sustainability_committee(self, config):
        """永續發展委員會組織架構說明 - 275字（含公司規模）"""
        company_context = self._get_company_context()
        # 取得年營收資訊（含阿拉伯數字）
        revenue_display = self.company_profile.get("revenue_display", "未知")
        prompt = f"""為永續發展委員會組織架構圖撰寫約275字說明文字，包含：委員會成立目的、組織架構重要性、召開會議的頻率、跨部門協作機制、制定政策與危機處理。語調專業。使用「我們」「本公司」等第一人稱表達。
本公司年營收約 {revenue_display}。
{company_context}
請根據企業規模調整描述，例如中小型企業可強調「精簡高效的組織架構」，中型企業可強調「完善的跨部門協作」。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_policy_description(self, config):
        """環境政策四大面向說明 - 275字"""
        prompt = """為環境政策四大面向圓餅圖撰寫約275字說明，包含：政策制定理念、論述四大面向風險監控、風險定義、風險評估和應對的意義與重要性，以及整體環境策略。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_tcfd_financial_disclosure(self, config):
        """4.3 TCFD 氣候財務揭露說明 - 約200字"""
        company_context = self._get_company_context()
        prompt = f"""為ESG報告撰寫約200字的TCFD氣候相關財務揭露說明，包含：
1. TCFD框架簡介與重要性
2. 氣候風險與機會識別方法
3. 公司採用TCFD的承諾與具體行動
4. 氣候風險對財務的潛在影響

語調專業，使用「本公司」「我們」等第一人稱表達。
{company_context}
請根據企業規模調整描述，例如中小型企業可強調「逐步建立TCFD管理機制」，中型企業可強調「完善的TCFD風險評估體系」。"""
        return self.generate(prompt, max_tokens=600)

    def generate_tcfd_matrix_analysis(self, config):
        """TCFD風險矩陣分析 - 275字"""
        prompt = """為TCFD風險矩陣圖撰寫約275字詳細說明，包含：矩陣圖用途與意義包括市場端客戶對永續產品的偏好上升，是品牌溢價機會、企業客戶或 B2B 對 ESG 要求提升，會導致更多合作與供應鏈整合機會，衝擊程度與可能性評估標準如成本壓力，氣候變遷導致能源成本增加、碳稅衝擊等優先級判斷機制、風險管控策略。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1200)

    def generate_ghg_calculation_method(self, config):
        """碳排放計算方法說明 - 275字"""
        prompt = """為碳排放數據表格撰寫約275字說明，包含：溫室氣體三大範疇定義、GHG Protocol計算標準、各範疇估算方法、數據收集以用電量乘上同業係數佔全體碳排九成以上，應符合GRI要求。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_electricity_policy(self, config):
        """電力使用與節能政策 - 275字"""
        prompt = """針對電力使用與節能政策撰寫約275字說明，包含：電力在碳排放中的重要性、一般節電措施、能源管理政策、電力使用效率提升策略。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_energy_efficiency_measures(self, config):
        """節能技術措施說明 - 275字（含投資預算建議）"""
        company_context = self._get_company_context()
        
        # 取得具體預算數字和年營收資訊（含阿拉伯數字）
        budget_display = self.company_profile.get("budget_display", "適當")
        revenue_display = self.company_profile.get("revenue_display", "未知")
        
        prompt = f"""為節能技術措施撰寫約275字說明，包含：LED智慧照明系統、智慧空調控制、建築通風優化、其他節能技術應用與效益。使用「我們」「本公司」等第一人稱表達。
本公司年營收約 {revenue_display}。
{company_context}
請在文中提及具體的投資計畫，例如「本公司預計投入約 {budget_display} 於節能設備更新」，並說明預期效益（如節電率、投資回收年限）。確保建議金額符合企業規模。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_green_planting_program(self, config):
        """綠色植栽計畫說明 - 275字"""
        prompt = """為綠色植栽計畫撰寫約275字說明，包含：SDGs生物多樣性目標、生態多元化重要性、森林認養復育計畫、綠色廠區植栽效益。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_water_management(self, config):
        """水資源管理說明 - 275字"""
        prompt = """為水資源管理撰寫約275字說明，包含：水資源環保重要性、省水計畫措施、用水效率提升、水資源循環利用策略。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_waste_management(self, config):
        """廢棄物管理說明 - 275字"""
        prompt = """為廢棄物管理撰寫約275字說明，包含：廢棄物分類處理、循環經濟理念、減廢措施、資源回收再利用流程。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_environmental_education(self, config):
        """環境教育與合作說明 - 275字"""
        prompt = """為環境教育與合作撰寫約275字說明，包含：學生教育營隊活動、員工家庭日環保宣導、濕地生態記錄計畫、社區環境合作項目。使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1000)

    def generate_sasb_analysis(self, config, industry, sasb_code, sasb_name):
        """SASB 產業分類分析 - ESG 專家洞察，從 26 個通用議題提出 5 個建議"""
        company_context = self._get_company_context()
        revenue_display = self.company_profile.get("revenue_display", "未知")
        
        prompt = f"""你是 ESG 專家。針對「{industry}」產業進行 SASB 分析，撰寫約350字說明。

【產業資訊】
- 產業名稱：{industry}
- SASB 代碼：{sasb_code}
- SASB 產業類別：{sasb_name}
- 本公司年營收約 {revenue_display}。
{company_context}

請以 ESG 專家的洞察角度，針對「{industry}」產業（SASB 分類：{sasb_name}）：
1. 分析該產業在 SASB 框架下的 ESG 重點議題
2. 從 SASB 通用議題中，針對此產業分類提出 5 個最相關的改善建議
3. 根據公司規模和產業特性，說明這些建議的優先順序和實施重點

請明確列出 5 個建議，並說明為何這些議題對「{industry}」產業特別重要。

語調專業，使用「我們」「本公司」等第一人稱表達。"""
        return self.generate(prompt, max_tokens=1800)