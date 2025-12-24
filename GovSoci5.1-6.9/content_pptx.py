"""Content engine for PPT generation (中文版)."""
import anthropic
import re
from typing import Optional, Dict, Any

from config_pptx import ANTHROPIC_API_KEY, CLAUDE_MODEL, CLAUDE_MODEL_FALLBACKS, LLM_WORD_COUNT
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

    # Slide-specific generators (中文版)
    def generate_governance_overview(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場、公司背景）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        company_context = self.env_context.get("company_context", "")
        
        prompt = f"你是{company_name}的{industry}產業 ESG 專家。"
        prompt += f"\n\n請撰寫約 420 字（對應 280 英文單字）的治理結構概述，用於 ESG 報告。"
        prompt += f"\n\n必須緊扣{industry}產業的治理特色，描述治理架構、董事會組成、監督節奏和利害關係人溝通。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用高階主管語調，保持專業且易於理解。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        if company_context:
            prompt += f"\n\n公司背景：{company_context}"
        
        return self._call(prompt, word_count=280, is_chinese=True)

    def generate_gender_equality_overview(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 市場）
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是一位{industry}產業人力資源專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 420 字（對應 280 英文單字）的性別平等倡議說明，用於 ESG 報告。"
        prompt += f"必須結合{industry}產業的特性，涵蓋領導層代表性、薪酬公平監控、包容性政策和可衡量目標。"
        prompt += "說明市場對多元包容的期望如何影響人才策略和政策制定。"
        prompt += "使用專業且具說服力的語調。"
        
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場趨勢背景：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=280, is_chinese=True)

    def generate_legal_alignment_overview(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場、TCFD 政策）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        
        prompt = f"你是{company_name}的{industry}產業 ESG 專家。"
        prompt += f"\n\n請撰寫約 420 字（對應 280 英文單字）的法規遵循說明，用於 ESG 報告。"
        prompt += f"\n\n必須緊扣{industry}產業的特性，說明如何維持法規對齊、協調合規利害關係人，以及升級法律風險。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用專業且清晰的語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n關鍵法規挑戰：{tcfd_policy[:300]}"
        
        return self._call(prompt, word_count=280, is_chinese=True)

    def generate_legal_appliance_overview(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場、TCFD 政策）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        
        prompt = f"你是{company_name}的{industry}產業 ESG 專家。"
        prompt += f"\n\n請撰寫約 420 字（對應 280 英文單字）的法律遵循計畫說明，用於 ESG 報告。"
        prompt += f"\n\n必須緊扣{industry}產業的特性，說明法律遵循計畫，強調架構、稽核軌跡、內部控制和對治理的益處。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用專業且具說服力的語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規遵循背景：{tcfd_policy[:300]}"
        
        return self._call(prompt, word_count=280, is_chinese=True)

    def generate_supervisory_board_overview(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 政策與市場）
        industry = self.env_context.get("industry", "該產業")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是一位{industry}產業治理專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 420 字（對應 280 英文單字）的稽核委員會概述，用於 ESG 報告。"
        prompt += f"必須說明{industry}產業稽核委員會的角色、委員會職責和報告節奏，支持長期治理韌性。"
        prompt += "特別強調稽核委員會如何監督氣候相關風險和市場轉型對治理的影響。"
        prompt += "使用專業且清晰的語調。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n關鍵法規挑戰：{tcfd_policy[:300]}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場轉型趨勢：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=280, is_chinese=True)

    # ------------------------------------------------------------------
    # Social chapter generators (Section 6.x) - 中文版
    # ------------------------------------------------------------------
    def generate_social_community_investment(self) -> str:
        # 整合環境段 log 資料（產業別、碳排放）
        industry = self.env_context.get("industry", "該產業")
        emission_context = self.env_context.get("emission_context", "")
        
        prompt = f"你是一位{industry}產業社會責任專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 360 字（對應 240 英文單字）的社區參與和社會投資議程說明，用於 ESG 報告。"
        prompt += "必須突出四大旗艦計畫（未來技能、健康、環境、動物福利）、資金邏輯，以及對齊 ISO 26000 條款 6.8.3 和 6.8.5。"
        prompt += "說明夥伴關係、志工服務和衡量框架如何確保包容性和韌性社區。"
        prompt += "使用專業且具說服力的語調。"
        
        if emission_context:
            prompt += f"\n\n環境績效背景：{emission_context}"
        
        return self._call(prompt, word_count=240, is_chinese=True)

    def generate_social_health_safety(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場、TCFD 政策）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        
        prompt = f"你是{company_name}的{industry}產業 ESG 專家。"
        prompt += f"\n\n請撰寫約 360 字（對應 240 英文單字）的員工健康、安全和福祉治理說明，用於 ESG 報告。"
        prompt += f"\n\n必須緊扣{industry}產業的特性，參考 ISO 45001 和 ISO 45003，涵蓋危害識別、心理健康資源、職業健康投資、參與機制和成功指標。"
        prompt += "說明氣候相關法規可能如何影響工作環境和職業安全規劃。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用專業且清晰的語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規變化背景：{tcfd_policy[:300]}"
        
        return self._call(prompt, word_count=240, is_chinese=True)

    def generate_social_diversity_policies(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 市場）
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是一位{industry}產業人力資源專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 345 字（對應 230 英文單字）的多元、包容和平等機會政策說明，用於 ESG 報告。"
        prompt += "必須說明性別多元計畫、同工同酬保證、反歧視培訓和導師計畫。"
        prompt += "強調治理節奏、領導問責和與人才策略的連結。"
        prompt += "說明市場對多元包容的期望如何推動多元包容政策。"
        prompt += "使用專業且具說服力的語調。"
        
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場趨勢背景：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_diversity_kpis(self) -> str:
        # 整合環境段 log 資料（產業別）
        industry = self.env_context.get("industry", "該產業")
        
        prompt = f"你是一位{industry}產業人力資源專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 345 字（對應 230 英文單字）的多元與包容關鍵績效指標說明，用於 ESG 報告。"
        prompt += "必須詳細說明女性管理層代表性、薪酬公平比率、培訓參與度、晉升速度和文化調查洞察等指標。"
        prompt += "加入一段關於內部多元與包容倡議的敘述，促進跨部門導師制度和弱勢群體的領導發展。"
        prompt += "使用專業且清晰的語調。"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_labor_rights(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 政策）
        industry = self.env_context.get("industry", "該產業")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        
        prompt = f"你是一位{industry}產業人權專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 345 字（對應 230 英文單字）的勞動權利方法說明，用於 ESG 報告。"
        prompt += "必須參考 ISO 26000 條款 6.3 關於人權的內容，說明組織的勞動權利方法。"
        prompt += "涵蓋盡職調查、申訴管道、供應商入職和對弱勢勞工群體的保護。"
        prompt += "說明法規變化可能如何影響勞動權利和就業實務。"
        prompt += "使用專業且具說服力的語調。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規變化背景：{tcfd_policy[:300]}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_fair_employment(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 政策）
        industry = self.env_context.get("industry", "該產業")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        
        prompt = f"你是一位{industry}產業勞動實務專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 345 字（對應 230 英文單字）的公平就業實務說明，用於 ESG 報告。"
        prompt += "必須對齊 ISO 26000 條款 6.4 關於勞動實務的內容，說明公平就業實務。"
        prompt += "涵蓋公平合約、工作時數、薪酬治理、社會對話和勞動力發展計畫。"
        prompt += "說明法規變化可能如何影響勞動權利和就業實務。"
        prompt += "使用專業且清晰的語調。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規變化背景：{tcfd_policy[:300]}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_action_plan_overview(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場、碳排放）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        emission_context = self.env_context.get("emission_context", "")
        
        prompt = f"你是{company_name}的{industry}產業 ESG 專家。"
        prompt += f"\n\n請撰寫約 330 字（對應 220 英文單字）的社會行動計畫概述，用於 ESG 報告。"
        prompt += f"\n\n必須緊扣{industry}產業的特性，為社會行動計畫表格提供敘述背景，概述優先順序邏輯、負責人、進度追蹤、利害關係人回饋循環，以及已完成倡議如何指導下一波承諾。"
        prompt += "說明市場期望和環境績效如何影響社會投資優先順序。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用專業且清晰的語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        if emission_context:
            prompt += f"\n\n環境績效背景：{emission_context}"
        
        return self._call(prompt, word_count=220, is_chinese=True)

    def generate_social_showcase_intro(self) -> str:
        # 整合環境段 log 資料（產業別、碳排放）
        industry = self.env_context.get("industry", "該產業")
        emission_context = self.env_context.get("emission_context", "")
        
        prompt = f"你是一位{industry}產業社會責任專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 270 字（對應 180 英文單字）的社區影響專案視覺展示介紹，用於 ESG 報告。"
        prompt += "必須介紹社區影響專案的視覺展示，總結為什麼未來技能計畫、環境清理活動和動物福利倡議對利害關係人重要，以及故事敘述、夥伴關係和衡量如何展現具體價值。"
        prompt += "使用專業且具說服力的語調。"
        
        if emission_context:
            prompt += f"\n\n環境影響背景：{emission_context}"
        
        return self._call(prompt, word_count=180, is_chinese=True)

    def generate_social_flow_explanation(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 市場）
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是一位{industry}產業社會影響專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 330 字（對應 220 英文單字）的社會影響邏輯鏈說明，用於 ESG 報告。"
        prompt += "必須說明從資源投入、活動、產出到長期社區價值的社會影響邏輯鏈。"
        prompt += "將流程連結到治理、學習循環，以及成果如何指導未來的資金優先順序。"
        prompt += "說明市場轉型如何影響社會影響的衡量和優先順序。"
        prompt += "使用專業且清晰的語調。"
        
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場轉型趨勢：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=220, is_chinese=True)

    def generate_social_product_responsibility(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是{company_name}的{industry}產業 ESG 專家。"
        prompt += f"\n\n請撰寫約 345 字（對應 230 英文單字）的產品責任計畫說明，用於 ESG 報告。"
        prompt += f"\n\n必須緊扣{industry}產業的特性，參考 ISO 9001、ISO 10377 和 ISO 26000 消費者議題（6.7），說明公司的產品責任計畫。"
        prompt += "涵蓋品質設計控制、客戶回饋分析、售後服務和對弱勢使用者的保護措施。"
        prompt += "說明市場對永續產品的期望如何影響產品責任策略。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用專業且具說服力的語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_customer_welfare(self) -> str:
        # 整合環境段 log 資料（公司名稱、產業、市場）
        company_name = self.env_context.get("company_name", "本公司")
        industry = self.env_context.get("industry", "該產業")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是{company_name}的{industry}產業 ESG 專家。"
        prompt += f"\n\n請撰寫約 345 字（對應 230 英文單字）的客戶福祉保護說明，用於 ESG 報告。"
        prompt += f"\n\n必須緊扣{industry}產業的特性，說明如何透過透明度、負責任行銷、資料隱私、申訴解決和生命週期管理來保護客戶福祉。"
        prompt += "包含與消費者保護機構的夥伴關係，以及用於監控信任的指標。"
        prompt += "說明市場對永續產品的期望如何影響客戶保護策略。"
        prompt += "\n\n使用「我們」和「本公司」，保持第一人稱視角，避免使用「貴公司」、「你們公司」等第三人稱。"
        prompt += "使用專業且清晰的語調。"
        
        if industry:
            prompt += f"\n\n產業別：{industry}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場摘要：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_innovation(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 政策與市場）
        industry = self.env_context.get("industry", "該產業")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是一位{industry}產業社會創新專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 345 字（對應 230 英文單字）的社會創新倡議說明，用於 ESG 報告。"
        prompt += "必須對齊 ISO 56000 和 ISO 26000 條款 6.8.9，總結社會創新倡議。"
        prompt += "聚焦包容性經濟參與、與社區共同創造，以及對社會企業的投資。"
        prompt += "說明市場轉型和法規變化如何推動社會創新。"
        prompt += "使用專業且具說服力的語調。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規變化背景：{tcfd_policy[:300]}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場轉型趨勢：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

    def generate_social_inclusive_economy(self) -> str:
        # 整合環境段 log 資料（產業別、TCFD 政策與市場）
        industry = self.env_context.get("industry", "該產業")
        tcfd_policy = self.env_context.get("tcfd_policy_context", "")
        tcfd_market = self.env_context.get("tcfd_market_context", "")
        
        prompt = f"你是一位{industry}產業包容經濟專家和 ESG 顧問。"
        prompt += f"請為一家{industry}公司撰寫約 345 字（對應 230 英文單字）的包容性經濟參與說明，用於 ESG 報告。"
        prompt += "必須說明如何透過供應商多元化、在地採購、影響力投資和共享價值夥伴關係來實現包容性經濟參與。"
        prompt += "強調社會回報和社區韌性成果的衡量。"
        prompt += "說明市場轉型和法規變化如何推動包容性經濟參與。"
        prompt += "使用專業且清晰的語調。"
        
        if tcfd_policy and len(tcfd_policy) < 500:
            prompt += f"\n\n法規變化背景：{tcfd_policy[:300]}"
        if tcfd_market and len(tcfd_market) < 500:
            prompt += f"\n\n市場轉型趨勢：{tcfd_market[:300]}"
        
        return self._call(prompt, word_count=230, is_chinese=True)

