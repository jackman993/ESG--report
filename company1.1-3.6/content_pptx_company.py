"""Extended content engine for PPT generation (includes company intro pages)."""
import anthropic
import re
import json
import os
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime

from config_pptx_company import (
    ANTHROPIC_API_KEY,
    CLAUDE_MODEL,
    CLAUDE_MODEL_FALLBACKS,
    LLM_WORD_COUNT,
)
from env_log_reader import load_latest_environment_log, get_prompt_context

LLM_WORD_COUNT = 280
# 中文約 1.5 字 = 1 英文單字，所以 280 英文單字約等於 420 中文字
CHINESE_CHAR_COUNT_MULTIPLIER = 1.5

META_PREFIXES = (
    "here is", "here's", "this response", "as an ai", "i am responding",
    "response in", "in english", "below is", "the following",
    # 中文的 meta 前綴
    "以下是", "這是", "作為", "我", "回應", "回答",
)


class PPTContentEngine:
    def __init__(self):
        if not ANTHROPIC_API_KEY:
            raise RuntimeError("ANTHROPIC_API_KEY is not configured.")
        self.client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        self.model = self._resolve_model()
        
        # 產業別：優先讀取，獨立管線，確保不被覆蓋
        self.industry = self._load_industry_directly()
        print(f"[產業別管線] 優先讀取: {repr(self.industry)}")
        
        # 載入環境段 log 資料（其他資料）
        self.env_log_data = self._load_environment_log()
        self.env_context = get_prompt_context(self.env_log_data)
        
        # 確保產業別不被覆蓋：如果 env_context 沒有，就用我們讀取的
        if not self.env_context.get("industry", "") and self.industry:
            self.env_context["industry"] = self.industry
            print(f"[產業別管線] 插入 env_context: {repr(self.industry)}")
        elif self.env_context.get("industry", ""):
            # 如果 env_context 有，但我們讀取的不同，以我們讀取的為準（優先）
            if self.industry and self.industry != self.env_context.get("industry", ""):
                print(f"[產業別管線] 覆蓋 env_context: {self.env_context.get('industry')} -> {self.industry}")
                self.env_context["industry"] = self.industry
            else:
                self.industry = self.env_context.get("industry", "")
    
    def _load_industry_directly(self) -> str:
        """
        產業 Express 通道：直接讀取最新的 Step 1 文件中的產業別
        不經過任何複雜流程，就這一個目的
        """
        log_dir = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")
        if not log_dir.exists():
            return ""
        
        # 找最新的 Step 1 文件
        for log_file in sorted(log_dir.glob("*.json"), key=lambda f: f.stat().st_mtime, reverse=True):
            try:
                with open(log_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                step = str(data.get("step", "")).lower()
                if "step 1" in step:
                    ind = data.get("industry", "")
                    if ind and str(ind).strip():
                        print(f"[產業 Express] 從 {log_file.name} 讀取: {ind}")
                        return str(ind).strip()
            except:
                continue
        return ""
    
    def _read_industry_analysis_directly(self) -> str:
        """
        直接從 log 文件讀取產業別分析（整段 150 字）
        完整解析 JSON，讀取整段 LLM 回應，不做正則抽取
        """
        log_dir = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")
        if not log_dir.exists():
            return ""
        
        # 優先找 industry_analysis.json 文件
        industry_analysis_files = sorted(
            log_dir.glob("session_*_industry_analysis.json"),
            key=lambda f: f.stat().st_mtime,
            reverse=True
        )
        
        # 如果沒有專門的 industry_analysis.json，則從所有 json 文件中找
        if not industry_analysis_files:
            all_json_files = sorted(
                log_dir.glob("*.json"),
                key=lambda f: f.stat().st_mtime,
                reverse=True
            )
        else:
            all_json_files = industry_analysis_files
        
        # 從最新的文件開始讀取
        for log_file in all_json_files:
            try:
                with open(log_file, "r", encoding="utf-8") as f:
                    data = json.load(f)  # 完整解析 JSON
                
                # 直接讀取 industry_analysis 字段（整段 150 字）
                industry_analysis = data.get("industry_analysis", "")
                
                if industry_analysis and isinstance(industry_analysis, str):
                    analysis_text = industry_analysis.strip()
                    if len(analysis_text) > 50:  # 至少 50 字才認為有效
                        print(f"[產業別分析] 從 {log_file.name} 讀取整段: {len(analysis_text)}字")
                        return analysis_text
            except Exception as e:
                print(f"[WARN] 讀取 {log_file.name} 失敗: {e}")
                continue
        
        return ""
    
    def _build_industry_first_prompt(self, base_prompt: str, industry: str) -> str:
        """
        產業 Express 通道：構建產業別優先的 prompt
        確保產業別在最前面，不被其他內容蓋過
        """
        if not industry:
            return base_prompt
        
        # 產業別優先：放在最前面
        industry_header = f"【產業別：{industry}】\n\n"
        industry_header += f"本公司在{industry}產業營運。所有內容必須緊扣{industry}產業的特性。\n\n"
        
        return industry_header + base_prompt

    def _resolve_model(self) -> str:
        candidates = []
        if CLAUDE_MODEL:
            candidates.append(CLAUDE_MODEL)
        for fallback in CLAUDE_MODEL_FALLBACKS:
            if fallback not in candidates:
                candidates.append(fallback)

        try:
            available = [m.id for m in self.client.models.list().data]
        except Exception:
            available = []

        for candidate in candidates:
            if candidate in available:
                return candidate
        return candidates[0] if candidates else "claude-3-haiku-20240307"

    def _format_expert_intro(self, company_name: str, industry: str) -> str:
        """
        格式化專家介紹語句，正確處理空產業別的情況
        
        Args:
            company_name: 公司名稱
            industry: 產業別（可能為空）
        
        Returns:
            格式化後的專家介紹語句
        """
        industry = industry.strip() if industry else ""
        if industry:
            return f"你是{company_name}的{industry}產業 ESG 專家。"
        else:
            return f"你是{company_name}的 ESG 專家。"

    def _call(self, prompt: str, word_count: int = LLM_WORD_COUNT, is_chinese: bool = True) -> str:
        """
        調用 LLM 生成內容
        
        Args:
            prompt: 提示詞
            word_count: 英文單字數（中文會自動轉換為字數）
            is_chinese: 是否為中文生成（預設 True）
        """
        if is_chinese:
            # 中文：將英文單字數轉換為中文字數
            char_count = int(word_count * CHINESE_CHAR_COUNT_MULTIPLIER)
            content = f"請用繁體中文撰寫約 {char_count} 字（對應 {word_count} 英文單字）。\n\n{prompt}"
        else:
            content = f"Respond in English with exactly {word_count} words.\n\n{prompt}"
        
        # 簡單檢查 prompt 是否包含產業別
        if "產業" in prompt:
            print(f"[OK] _call: prompt 包含產業別")
        
        response = self.client.messages.create(
            model=self.model,
            max_tokens=word_count * 5,
            messages=[
                {
                    "role": "user",
                    "content": content,
                }
            ],
        )
        text = response.content[0].text if response.content else ""
        return self._clean(text, is_chinese=is_chinese)

    @staticmethod
    def _clean(text: str, is_chinese: bool = True) -> str:
        """
        清理 LLM 回應文字
        
        Args:
            text: 原始文字
            is_chinese: 是否為中文（影響字數計算方式）
        """
        lines = []
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            lower = line.lower()
            if any(lower.startswith(prefix) for prefix in META_PREFIXES):
                continue
            lines.append(line)
        
        if is_chinese:
            # 中文：直接連接，按字數計算
            cleaned = "".join(lines)
            # 移除 meta 註解
            cleaned = re.sub(r"^\[?約?\s*\d+\s*字\]?[：:\-]*\s*", "", cleaned, flags=re.IGNORECASE)
            cleaned = re.sub(r"^\[?exactly\s+\d+\s+words\]?[:\-]*\s*", "", cleaned, flags=re.IGNORECASE)
            # 中文不需要分割單字，直接返回
            return cleaned.strip()
        else:
            # 英文：按單字計算
            cleaned = " ".join(lines)
            cleaned = re.sub(r"^\[?exactly\s+\d+\s+words\]?[:\-]*\s*", "", cleaned, flags=re.IGNORECASE)
            words = cleaned.split()
            if not words:
                return ""
            if len(words) > LLM_WORD_COUNT:
                words = words[:LLM_WORD_COUNT]
            elif len(words) < LLM_WORD_COUNT:
                words.extend([" "] * (LLM_WORD_COUNT - len(words)))
            return " ".join(words).strip()

    def _load_environment_log(self) -> Optional[Dict[str, Any]]:
        """
        載入最新的環境段 log 資料
        
        Returns:
            Log 資料 dict，如果找不到則返回 None
        """
        return load_latest_environment_log()

    # --- Governance section generators (unchanged) ---
    def generate_governance_overview(self) -> str:
        # 整合環境段 log 資料（產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "").strip()
        
        # 調試信息
        print(f"[DEBUG] generate_governance_overview: industry={repr(industry)}, company_name={repr(company_name)}")
        
        if not industry:
            print(f"[ERROR] generate_governance_overview: 產業別為空！env_context keys: {list(self.env_context.keys())}")
            print(f"[DEBUG] env_log_data industry: {self.env_log_data.get('industry', 'NOT FOUND') if self.env_log_data else 'env_log_data is None'}")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請撰寫高階主管級別的敘述，描述治理結構、董事會組成、監督節奏和利害關係人溝通。"
        
        if industry:
            prompt += f"\n\n【重要】本公司在{industry}產業營運。請根據{industry}產業的特性，分析關係人、相關法律合規、市場衝擊，並在內容中明確提及{industry}產業相關的法規、風險和治理要求。"
            prompt += f"內容必須緊扣{industry}產業的特性，確保與{industry}產業緊密相關。"
        else:
            prompt += f"\n\n【重要】請根據本公司所屬產業的特性，分析關係人、相關法律合規、市場衝擊，並在內容中明確提及與產業相關的法規、風險和治理要求。"
        
        return self._call(prompt, is_chinese=True)

    def generate_gender_equality_overview(self) -> str:
        return self._call(
            "Describe gender equality initiatives covering leadership representation, pay equity monitoring, inclusive policies, and measurable goals."
        )

    def generate_legal_alignment_overview(self) -> str:
        # 整合環境段 log 資料（產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請總結公司如何維持與法規的對齊、協調合規利害關係人，以及升級法律風險。"
        
        if industry:
            prompt += f"\n\n【重要】本公司在{industry}產業營運。請根據{industry}產業的特性，分析相關法律合規，並說明如何維持與{industry}產業特定的環境法規、財務合規要求的對齊。"
        else:
            prompt += f"\n\n【重要】請根據本公司所屬產業的特性，分析相關法律合規，並說明如何維持與產業特定的環境法規、財務合規要求的對齊。"
        
        return self._call(prompt, is_chinese=True)

    def generate_legal_appliance_overview(self) -> str:
        # 整合環境段 log 資料（產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請說明法律遵循計畫，強調框架、審計追蹤、內部控制和對治理的益處。"
        
        if industry:
            prompt += f"\n\n【重要】本公司在{industry}產業營運。請根據{industry}產業的特性，分析相關法律合規，並說明法律遵循計畫如何應對{industry}產業特定的環境法規、財務合規框架、審計追蹤和內部控制要求。"
        else:
            prompt += f"\n\n【重要】請根據本公司所屬產業的特性，分析相關法律合規，並說明法律遵循計畫如何應對產業特定的環境法規、財務合規框架、審計追蹤和內部控制要求。"
        
        return self._call(prompt, is_chinese=True)

    def generate_supervisory_board_overview(self) -> str:
        return self._call(
            "Outline the supervisory board's roles, committee responsibilities, and reporting rhythm supporting long-term governance resilience."
        )

    # --- Social section generators (unchanged) ---
    def generate_social_community_investment(self) -> str:
        return self._call(
            "Describe the company's community engagement and social investment agenda highlighting the four flagship programmes (Future Skills, Health, Environment, Animal Welfare), funding logic, and alignment with ISO 26000 clauses 6.8.3 and 6.8.5. "
            "Explain how partnerships, volunteerism, and measurement frameworks ensure inclusive and resilient communities.",
            word_count=240,
        )

    def generate_social_health_safety(self) -> str:
        # 整合環境段 log 資料（產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請總結員工健康、安全和福祉治理，參考 ISO 45001 和 ISO 45003。"
        prompt += "涵蓋危害識別、心理健康資源、職業健康投資、參與機制和成功指標。"
        
        if industry:
            prompt += f"\n\n【重要】本公司在{industry}產業營運。請根據{industry}產業的特性，分析員工勞安衛（勞動安全衛生），並說明{industry}產業特定的職業健康與安全要求、危害識別流程、安全協議和適用於{industry}產業工作者的健康法規。"
        else:
            prompt += f"\n\n【重要】請根據本公司所屬產業的特性，分析員工勞安衛（勞動安全衛生），並說明產業特定的職業健康與安全要求、危害識別流程、安全協議和適用於產業工作者的健康法規。"
        
        return self._call(prompt, word_count=240, is_chinese=True)

    def generate_social_diversity_policies(self) -> str:
        return self._call(
            "Explain diversity, inclusion, and equal opportunity policies including the Gender Diversity Plan, Equal Pay for Equal Work assurance, anti-discrimination training, and mentorship programmes. "
            "Highlight governance cadence, leadership accountability, and links to talent strategy.",
            word_count=230,
        )

    def generate_social_diversity_kpis(self) -> str:
        return self._call(
            "Detail key performance indicators for diversity and inclusion such as female management representation, pay equity ratios, training participation, promotion velocity, and culture survey insights. "
            "Add a short narrative on the internal D&I initiative that promotes cross-department mentorship and leadership development for underrepresented groups.",
            word_count=230,
        )

    def generate_social_labor_rights(self) -> str:
        return self._call(
            "Describe the organisation's labour rights approach referencing ISO 26000 clause 6.3 on human rights. "
            "Include due diligence, grievance channels, supplier onboarding, and protections for vulnerable worker groups.",
            word_count=230,
        )

    def generate_social_fair_employment(self) -> str:
        return self._call(
            "Explain fair employment practices in line with ISO 26000 clause 6.4 on labour practices, covering fair contracts, working hours, compensation governance, social dialogue, and workforce development programmes.",
            word_count=230,
        )

    def generate_social_action_plan_overview(self) -> str:
        # 整合環境段 log 資料（產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請為社會行動計畫表格提供敘述性背景，概述優先順序邏輯、負責人、進度追蹤、利害關係人回饋循環，以及已完成的倡議如何為下一波承諾提供資訊。"
        
        if industry:
            prompt += f"\n\n【重要】本公司在{industry}產業營運。請根據{industry}產業的特性，分析關係人、市場衝擊，並概述社會行動計畫如何應對與{industry}產業相關的環境影響和社會責任，包括{industry}產業特定的環境挑戰和緩解策略。"
        else:
            prompt += f"\n\n【重要】請根據本公司所屬產業的特性，分析關係人、市場衝擊，並概述社會行動計畫如何應對與產業相關的環境影響和社會責任，包括產業特定的環境挑戰和緩解策略。"
        
        return self._call(prompt, word_count=220, is_chinese=True)

    def generate_social_showcase_intro(self) -> str:
        return self._call(
            "Introduce the visual showcase of community impact projects. Summarise why the Future Skills programme, environmental clean-up campaigns, and animal welfare initiatives matter to stakeholders and how storytelling, partnerships, and measurement demonstrate tangible value.",
            word_count=180,
        )

    def generate_social_flow_explanation(self) -> str:
        return self._call(
            "Explain the social impact logic chain from resource inputs through activities, outputs, and longer-term community value. "
            "Connect the flow to governance, learning loops, and how outcomes inform future funding priorities.",
            word_count=220,
        )

    def generate_social_product_responsibility(self) -> str:
        # 整合環境段 log 資料（產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請描述公司的產品責任計畫，參考 ISO 9001、ISO 10377 和 ISO 26000 消費者議題（6.7）。"
        prompt += "涵蓋品質設計控制、客戶回饋分析、售後服務和對弱勢使用者的保護措施。"
        prompt += f"\n\n【重要】請根據本公司產業，分析關係人、相關法律合規、市場衝擊，並描述產品責任計畫如何應對產業特定的環境影響、產品生命週期考量和適用於產業的產品安全要求。"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_customer_welfare(self) -> str:
        # 整合環境段 log 資料（產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請說明如何透過透明度、負責任的行銷、資料隱私、投訴解決和生命週期管理來保護客戶福祉。"
        prompt += "包括與消費者保護機構的合作夥伴關係和用於監控信任的指標。"
        prompt += f"\n\n【重要】請根據本公司產業，分析關係人、市場衝擊，並說明客戶福祉保護措施如何應對產業產品/服務的環境影響、產品責任考量和適用於產業的客戶安全要求。"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_innovation(self) -> str:
        return self._call(
            "Summarise social innovation initiatives aligned with ISO 56000 and ISO 26000 clause 6.8.9, focusing on inclusive economic participation, co-creation with communities, and investment in social enterprises.",
            word_count=230,
        )

    def generate_social_inclusive_economy(self) -> str:
        return self._call(
            "Explain how inclusive economic participation is enabled through supplier diversity, local procurement, impact investing, and shared value partnerships. "
            "Highlight measurement of social return and community resilience outcomes.",
            word_count=230,
        )

    # --- Company section additions (中文版) ---
    def generate_ceo_message(self) -> str:
        # 整合環境段 log 資料
        company_context = self.env_context.get("company_context", "")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        
        prompt = "請撰寫約 330 字（對應 220 英文單字）的 CEO 訊息，用於 ESG 永續報告書。"
        prompt += "平衡啟發性與責任感：歡迎讀者、概述永續願景、強調近期 ESG 成就、"
        prompt += "承認仍面臨的挑戰，並呼籲利害關係人共同參與轉型旅程。"
        prompt += "使用「我們」和「本公司」，保持溫暖且專業的語調，避免元評論。"
        
        if company_context:
            prompt += f"\n\n公司背景：{company_context}"
        if tcfd_policy and len(tcfd_policy) < 500:  # 避免 prompt 過長
            prompt += f"\n\n關鍵法規挑戰：{tcfd_policy[:300]}"
        
        return self._call(prompt, word_count=220, is_chinese=True)

    def generate_cooperation_info(self) -> str:
        # 直接從 log 文件讀取產業別分析（不經過 JSON 解析和多層轉換）
        industry = self.industry
        company_name = self.env_context.get("company_name", "本公司") if self.env_context else "本公司"
        company_context = self.env_context.get("company_context", "") if self.env_context else ""
        tcfd_market = self.env_context.get("tcfd_market_context", "") if self.env_context else ""
        
        # 直接讀取 log 文件中的產業別分析（150字）
        industry_analysis = self._read_industry_analysis_directly()
        
        # 調試：檢查實際值
        print(f"[DEBUG generate_cooperation_info] industry={repr(industry)}")
        print(f"[DEBUG generate_cooperation_info] industry_analysis 長度={len(industry_analysis) if industry_analysis else 0}")
        print(f"[DEBUG generate_cooperation_info] industry_analysis 前100字={industry_analysis[:100] if industry_analysis else 'None'}")
        
        # 產業別優先：放在最前面，不能被其他內容蓋過
        if industry:
            prompt = f"【產業別：{industry}】\n\n"
            prompt += f"本公司在{industry}產業營運。所有內容必須緊扣{industry}產業的特性。\n\n"
            
            # 直接插入產業別分析（150字，不經過解析）
            if industry_analysis:
                prompt += f"【產業別分析】（這是所有分析的基礎，必須遵守）：\n{industry_analysis}\n\n"
                print(f"[OK] 已直接插入產業別分析到 prompt（{len(industry_analysis)}字）")
            else:
                print(f"[WARN] industry_analysis 為空，無法加入")
            
            prompt += self._format_expert_intro(company_name, industry)
            prompt += f"\n\n請撰寫約 345 字（對應 230 英文單字）描述公司的合作概況，用於 ESG 報告。"
            prompt += f"\n\n【必須遵守】在第一句中使用 {{COMPANY_NAME}} 作為公司名稱的佔位符，並明確提及{industry}產業。"
            prompt += f"例如，以「{{COMPANY_NAME}} 是一家{industry}公司，擁有豐富的歷史...」或「{{COMPANY_NAME}} 在{industry}產業深耕多年...」開頭。"
            prompt += f"\n\n【必須遵守】請根據{industry}產業的特性，分析關係人、相關法律合規、市場衝擊，說明{industry}產業面臨的主要 ESG 挑戰（如環境影響、社會責任、治理需求），以及公司如何透過合作夥伴關係、組織架構和策略規劃來應對這些挑戰。"
            prompt += f"內容必須緊扣{industry}產業的特性，明確提及{industry}產業相關的法規、風險和治理要求。"
        else:
            prompt = self._format_expert_intro(company_name, industry)
            prompt += f"\n\n請撰寫約 345 字（對應 230 英文單字）描述公司的合作概況，用於 ESG 報告。"
            prompt += "\n\n重要：在第一句中使用 {COMPANY_NAME} 作為公司名稱的佔位符。"
            prompt += "例如，以「{COMPANY_NAME} 公司擁有豐富的歷史...」或「{COMPANY_NAME} 是一家多元化...」開頭。"
            prompt += f"\n\n【重要】請根據本公司所屬產業的特性，分析關係人、相關法律合規、市場衝擊，說明產業面臨的主要 ESG 挑戰（如環境影響、社會責任、治理需求），以及公司如何透過合作夥伴關係、組織架構和策略規劃來應對這些挑戰。"
            if company_context:
                prompt += f"\n\n公司背景：{company_context}"
        
        prompt += "\n\n總結背景、商業模式、地理足跡、策略夥伴關係和組織架構，強調使命、價值觀，以及合作如何支撐長期競爭力。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用簡潔的中文，不使用項目符號，保持高階主管語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        
        # 最後檢查 prompt 是否包含產業別
        print(f"\n{'='*60}")
        print(f"[檢查] industry 變數值: {repr(industry)}")
        print(f"[檢查] industry 是否為空: {not industry}")
        print(f"[檢查] prompt 長度: {len(prompt)}")
        print(f"[檢查] prompt 是否包含產業別 '{industry}': {industry in prompt if industry else 'N/A'}")
        if industry:
            print(f"[檢查] prompt 中包含 '{industry}' 的次數: {prompt.count(industry)}")
        print(f"[檢查] prompt 完整內容:")
        print(prompt)
        print(f"{'='*60}\n")
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_cooperation_financial(self) -> str:
        # 整合環境段 log 資料（營收資訊）
        company_context = self.env_context.get("company_context", "")
        
        prompt = "請撰寫約 345 字（對應 230 英文單字）詳細說明財務穩定性、營收成長、資本投資和法規遵循。"
        prompt += "討論風險管理、流動性紀律、再投資優先順序和保證機制。"
        prompt += "保持自信且透明的語調，避免元評論，以單一流暢段落呈現敘述。"
        
        if company_context:
            prompt += f"\n\n公司財務背景：{company_context}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_stakeholder_identify(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請撰寫約 375 字（對應 250 英文單字）的利害關係人識別章節，用於 ESG 報告。"
        prompt += f"\n\n【重要】請根據本公司產業，分析關係人，識別關鍵利害關係人群體，如投資人、客戶、員工、監管機構、供應商和社區，說明每個群體的重要性。"
        prompt += "討論他們的期望、若被忽視的潛在風險，以及如何設定優先順序。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用簡潔的高階主管語調，不使用項目符號或標題。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

    def generate_stakeholder_analysis(self) -> str:
        # 整合環境段 log 資料（市場趨勢、產業別）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請撰寫約 375 字（對應 250 英文單字）分析利害關係人的需求和影響力。"
        prompt += "涵蓋顯著性（權力、合法性、緊迫性）、重大議題、溝通節奏，以及監控參與成效的關鍵績效指標。"
        prompt += "保持高階主管語調，使用簡潔的中文，避免項目符號。"
        prompt += f"\n\n【重要】請根據本公司產業，分析關係人，說明產業中哪些利害關係人群體最具影響力，以及他們優先關注的議題，並討論產業特定的重大議題、溝通方式和參與指標。"
        
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場趨勢背景：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

    def generate_material_issues_text(self) -> str:
        # 整合環境段 log 資料（TCFD 政策與市場）
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = "請撰寫約 375 字（對應 250 英文單字）總結公司的重大議題概況。"
        prompt += "說明優先順序設定方法、與策略和風險的連結、利害關係人意見，以及預期成果。"
        prompt += "保持分析性且易於理解的語調，不使用元評論或項目列表。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規與政策風險：{tcfd_policy[:250]}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場趨勢風險：{tcfd_market[:250]}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

    def generate_materiality_summary(self) -> str:
        # 整合環境段 log 資料（TCFD 政策與市場）
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = "請撰寫約 375 字（對應 250 英文單字）總結重大性評估方法和主要發現。"
        prompt += "涵蓋雙重重大性（影響與財務）、利害關係人參與、矩陣解讀，以及由此產生的行動。"
        prompt += "使用清晰的文字，以單一連貫的敘述呈現。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規與政策風險：{tcfd_policy[:250]}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場趨勢風險：{tcfd_market[:250]}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

    def generate_sustainability_strategy_intro(self) -> str:
        return self._call(
            "請撰寫約 165 字（對應 110 英文單字）介紹永續策略與行動計畫。"
            "說明三大策略支柱、如何與營運整合，以及附表中描述的未來方向。"
            "使用「我們」和「本公司」，保持高階主管語調，不使用元評論。",
            word_count=110,
            is_chinese=True
        )

    def generate_esg_pillars(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場、碳排放）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        emission_context = self.env_context.get("emission_context", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請撰寫約 375 字（對應 250 英文單字）說明公司的 ESG 核心支柱，涵蓋地球（Planet）、產品（Products）和人員（People），用於 ESG 報告。"
        prompt += f"\n\n【重要】請根據本公司產業，分析關係人、相關法律合規、市場衝擊，討論重點領域、跨價值鏈的整合、衡量實務和問責機制。"
        prompt += "提及表格顯示每個支柱的摘要和倡議。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用簡潔的中文，避免項目符號。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        if emission_context:
            prompt += f"\n\n環境績效：{emission_context}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

    def generate_esg_roadmap_context(self) -> str:
        # 整合環境段 log 資料（TCFD 政策與碳排放）
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        emission_context = self.env_context.get("emission_context", "")
        
        prompt = "請撰寫約 345 字（對應 230 英文單字）敘述 ESG 路線圖。"
        prompt += "總結基礎承諾如何演進為風險整合、營運脫碳和前瞻性合規。"
        prompt += "強調里程碑、治理負責人，以及路線圖如何指導投資和利害關係人參與。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n關鍵法規挑戰：{tcfd_policy[:300]}"
        if emission_context:
            prompt += f"\n\n環境績效現況：{emission_context}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_stakeholder_communication(self) -> str:
        # 整合環境段 log 資料（市場趨勢）
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = "請提供約 375 字（對應 250 英文單字）描述利害關係人溝通實務，以補充重大性雷達圖。"
        prompt += "涵蓋參與節奏、回饋管道、揭露承諾、升級協議，以及洞察如何影響策略決策。"
        
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場趨勢背景：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

    def generate_sdg_summary(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場、碳排放與公司背景）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        emission_context = self.env_context.get("emission_context", "")
        company_context = self.env_context.get("company_context", "")
        
        prompt = self._format_expert_intro(company_name, industry)
        prompt += f"\n\n請撰寫約 375 字（對應 250 英文單字）說明公司如何將永續目標與標示的 SDG 圖示對齊，用於 ESG 報告。"
        prompt += f"\n\n【重要】請根據本公司產業，分析關係人，說明哪些 SDG 與產業最相關，以及產業中利害關係人的關注如何與特定 SDG 目標對齊。"
        prompt += f"詳細說明產業特定的計畫、夥伴關係和指標，這些計畫、夥伴關係和指標解決產業背景下的利害關係人優先事項。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用簡潔的中文，保持專業且易於理解的語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        if emission_context:
            prompt += f"\n\n環境績效：{emission_context}"
        if company_context:
            prompt += f"\n\n公司背景：{company_context}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

    def generate_risk_management_overview(self) -> str:
        # 整合環境段 log 資料（TCFD 政策風險）
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        
        prompt = "請撰寫約 375 字（對應 250 英文單字）概述與重大 ESG 議題相關的企業風險管理方法。"
        prompt += "涵蓋治理結構、評估節奏、緩解規劃、控制監控、升級和保證。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n關鍵氣候相關法規風險：{tcfd_policy[:300]}"
        
        return self._call(prompt, word_count=250, is_chinese=True)

