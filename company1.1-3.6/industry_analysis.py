"""
產業別分析生成器
在 Step 1 用戶按「生成 TCFD」時，第一個 LLM 調用生成產業別分析
"""
import anthropic
import json
from pathlib import Path
from typing import Dict, Any
from datetime import datetime

from config_pptx_company import ANTHROPIC_API_KEY, CLAUDE_MODEL, CLAUDE_MODEL_FALLBACKS

# 預設 log 路徑
DEFAULT_LOG_DIR = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")


def generate_industry_analysis(industry: str, monthly_electricity_bill_ntd: float, session_id: str) -> Dict[str, Any]:
    """
    生成產業別分析（150字）
    這是所有 LLM 調用的第一個，生成基礎數據
    
    Args:
        industry: 產業別
        monthly_electricity_bill_ntd: 月電費（NTD）
        session_id: Session ID
    
    Returns:
        包含產業別分析的 dict
    """
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY is not configured.")
    
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    
    # 解析模型
    model = CLAUDE_MODEL if CLAUDE_MODEL else "claude-3-haiku-20240307"
    
    # 構建 prompt
    prompt = f"""請根據以下資訊，撰寫約 150 字的產業別分析：

產業別：{industry}
月電費：{monthly_electricity_bill_ntd:,.0f} NTD

請分析以下內容：
1. 產業別相關規範要求
2. 市場趨勢
3. 風險分析
4. 判斷耗能等級（高耗能/中耗能/低耗能）
5. 根據耗能等級計算年營收：
   - 高耗能：月電費 × 300倍
   - 中耗能：月電費 × 600倍
   - 低耗能：月電費 × 1200倍

請用繁體中文撰寫，並在最後明確標示：
- 耗能等級：[高耗能/中耗能/低耗能]
- 計算年營收：[數值] NTD（不顯示計算過程，只顯示結果）

格式範例：
「{industry}產業面臨...規範要求...市場趨勢...風險分析...耗能等級：中耗能。計算年營收：54,000,000 NTD。」"""

    # 調用 LLM
    response = client.messages.create(
        model=model,
        max_tokens=500,
        messages=[
            {
                "role": "user",
                "content": prompt,
            }
        ],
    )
    
    analysis_text = response.content[0].text if response.content else ""
    
    # 解析結果
    energy_level = "中耗能"  # 預設值
    annual_revenue_ntd = monthly_electricity_bill_ntd * 600  # 預設中耗能
    
    # 從分析文字中提取耗能等級和年營收
    if "高耗能" in analysis_text:
        energy_level = "高耗能"
        annual_revenue_ntd = monthly_electricity_bill_ntd * 300
    elif "中耗能" in analysis_text:
        energy_level = "中耗能"
        annual_revenue_ntd = monthly_electricity_bill_ntd * 600
    elif "低耗能" in analysis_text:
        energy_level = "低耗能"
        annual_revenue_ntd = monthly_electricity_bill_ntd * 1200
    
    # 嘗試從文字中提取年營收數值（如果 LLM 有明確標示）
    import re
    revenue_match = re.search(r'計算年營收[：:]\s*([\d,]+)', analysis_text)
    if revenue_match:
        try:
            annual_revenue_ntd = float(revenue_match.group(1).replace(',', ''))
        except:
            pass
    
    # 構建結果
    result = {
        "industry": industry,
        "monthly_electricity_bill_ntd": monthly_electricity_bill_ntd,
        "industry_analysis": analysis_text.strip(),
        "energy_level": energy_level,
        "estimated_annual_revenue_ntd": annual_revenue_ntd,
        "estimated_annual_revenue_display": f"{annual_revenue_ntd/10000:.0f}萬元 ({annual_revenue_ntd:,.0f} NTD)",
        "session_id": session_id,
        "timestamp": datetime.now().isoformat(),
        "step": "Step 1 - 產業別分析"
    }
    
    # 寫入 log
    save_industry_analysis_to_log(result)
    
    return result


def save_industry_analysis_to_log(data: Dict[str, Any]) -> None:
    """
    將產業別分析硬寫入 log 文件
    """
    log_dir = DEFAULT_LOG_DIR
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # 使用 session_id 作為文件名
    session_id = data.get("session_id", datetime.now().strftime("%Y%m%d_%H%M%S"))
    log_file = log_dir / f"session_{session_id}_industry_analysis.json"
    
    # 如果已有同 session 的文件，先讀取並合併
    if log_file.exists():
        try:
            with open(log_file, "r", encoding="utf-8") as f:
                existing_data = json.load(f)
            # 合併數據（產業別分析優先）
            existing_data.update(data)
            data = existing_data
        except:
            pass
    
    # 寫入文件
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"[產業別分析] 已寫入 log: {log_file.name}")

