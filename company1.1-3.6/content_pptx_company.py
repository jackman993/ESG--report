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
        # 載入環境段 log 資料
        self.env_log_data = self._load_environment_log()
        self.env_context = get_prompt_context(self.env_log_data)

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
        return self._call(
            "Create executive-level narration describing the governance structure, board composition, oversight cadence, and stakeholder communication."
        )

    def generate_gender_equality_overview(self) -> str:
        return self._call(
            "Describe gender equality initiatives covering leadership representation, pay equity monitoring, inclusive policies, and measurable goals."
        )

    def generate_legal_alignment_overview(self) -> str:
        return self._call(
            "Summarise how the company maintains alignment with regulations, coordinates compliance stakeholders, and escalates legal risks."
        )

    def generate_legal_appliance_overview(self) -> str:
        return self._call(
            "Explain the legal appliance program, highlighting frameworks, audit trails, internal controls, and benefits for governance."
        )

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
        return self._call(
            "Summarise employee health, safety, and wellbeing governance referencing ISO 45001 and ISO 45003. "
            "Cover hazard identification, mental health resources, occupational health investments, participation mechanisms, and success indicators.",
            word_count=240,
        )

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
        return self._call(
            "Provide narrative context for the social action plan table, outlining prioritisation logic, accountable owners, progress tracking, stakeholder feedback loops, and how completed initiatives inform the next wave of commitments.",
            word_count=220,
        )

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
        return self._call(
            "Describe the company's product responsibility programme referencing ISO 9001, ISO 10377, and ISO 26000 consumer issues (6.7). "
            "Cover quality design controls, customer feedback analytics, after-sales care, and safeguards for vulnerable users.",
            word_count=230,
        )

    def generate_social_customer_welfare(self) -> str:
        return self._call(
            "Explain how customer welfare is protected through transparency, responsible marketing, data privacy, complaint resolution, and lifecycle stewardship. "
            "Include partnerships with consumer protection bodies and indicators used to monitor trust.",
            word_count=230,
        )

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
        # 整合環境段 log 資料（公司背景）
        company_context = self.env_context.get("company_context", "")
        
        prompt = "請提供約 345 字（對應 230 英文單字）描述公司的合作概況。"
        prompt += "重要：在第一句中使用 {COMPANY_NAME} 作為公司名稱的佔位符。"
        prompt += "例如，以「{COMPANY_NAME} 公司擁有豐富的歷史...」或「{COMPANY_NAME} 是一家多元化...」開頭。"
        prompt += "總結背景、商業模式、地理足跡、策略夥伴關係和組織架構。"
        prompt += "強調使命、價值觀，以及合作如何支撐長期競爭力。"
        prompt += "使用簡潔的中文，不使用項目符號，保持高階主管語調。"
        
        if company_context:
            prompt += f"\n\n公司背景：{company_context}"
        
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
        return self._call(
            "請撰寫約 375 字（對應 250 英文單字）的利害關係人識別章節，用於 ESG 報告。"
            "識別關鍵利害關係人群體，如投資人、客戶、員工、監管機構、供應商和社區，說明每個群體的重要性。"
            "討論他們的期望、若被忽視的潛在風險，以及如何設定優先順序。"
            "使用簡潔的高階主管語調，不使用項目符號或標題。",
            word_count=250,
            is_chinese=True
        )

    def generate_stakeholder_analysis(self) -> str:
        # 整合環境段 log 資料（市場趨勢）
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = "請撰寫約 375 字（對應 250 英文單字）分析利害關係人的需求和影響力。"
        prompt += "涵蓋顯著性（權力、合法性、緊迫性）、重大議題、溝通節奏，以及監控參與成效的關鍵績效指標。"
        prompt += "保持高階主管語調，使用簡潔的中文，避免項目符號。"
        
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
        # 整合環境段 log 資料（碳排放）
        emission_context = self.env_context.get("emission_context", "")
        
        prompt = "請撰寫約 375 字（對應 250 英文單字）說明公司的 ESG 核心支柱，涵蓋地球（Planet）、產品（Products）和人員（People）。"
        prompt += "討論重點領域、跨價值鏈的整合、衡量實務和問責機制。"
        prompt += "提及表格顯示每個支柱的摘要和倡議。"
        prompt += "使用簡潔的中文，避免項目符號。"
        
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
        # 整合環境段 log 資料（碳排放與公司背景）
        emission_context = self.env_context.get("emission_context", "")
        company_context = self.env_context.get("company_context", "")
        
        prompt = "請撰寫約 375 字（對應 250 英文單字）說明公司如何將永續目標與標示的 SDG 圖示對齊。"
        prompt += "詳細說明關鍵計畫、夥伴關係、指標和治理結構，這些將 SDG 與商業價值和利害關係人期望連結。"
        
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

