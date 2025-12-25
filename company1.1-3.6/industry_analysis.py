"""
產業別分析生成器
在 Step 1 用戶按「生成 TCFD」時，第一個 LLM 調用生成產業別分析
"""
import anthropic
import json
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

from config_pptx_company import ANTHROPIC_API_KEY, CLAUDE_MODEL, CLAUDE_MODEL_FALLBACKS

# 預設 log 路徑
DEFAULT_LOG_DIR = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")


def _load_emission_data_from_log(session_id: str) -> Optional[float]:
    """
    從前階段的 log 讀取碳排數據
    """
    log_dir = DEFAULT_LOG_DIR
    if not log_dir.exists():
        return None
    
    # 找同 session 的 log 文件
    for log_file in sorted(log_dir.glob("*.json"), key=lambda f: f.stat().st_mtime, reverse=True):
        try:
            with open(log_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            # 檢查是否為同一個 session
            file_session_id = data.get("session_id", "")
            if str(file_session_id) != str(session_id):
                continue
            
            # 優先從 emission_result 讀取
            emission_result = data.get("emission_result", {})
            if emission_result and "total" in emission_result:
                total = emission_result.get("total", 0.0)
                if total and total > 0:
                    print(f"[產業別分析] 從 {log_file.name} 讀取碳排數據: {total:.2f} tCO₂e")
                    return float(total)
            
            # 從 emission_data 讀取
            emission_data = data.get("emission_data", {})
            if emission_data and "total" in emission_data:
                total = emission_data.get("total", 0.0)
                if total and total > 0:
                    print(f"[產業別分析] 從 {log_file.name} 讀取碳排數據: {total:.2f} tCO₂e")
                    return float(total)
        except Exception as e:
            continue
    
    return None


def generate_industry_analysis(industry: str, monthly_electricity_bill_ntd: float, session_id: str) -> Dict[str, Any]:
    """
    生成產業別分析（150字）
    這是所有 LLM 調用的第一個，生成基礎數據
    會從前階段的 log 讀取碳排數據，確保分析包含具體數據
    
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
    
    # 從前階段 log 讀取碳排數據
    emission_total_tco2e = _load_emission_data_from_log(session_id)
    
    # 構建 prompt（包含碳排數據）
    prompt = f"""請根據以下資訊，撰寫約 150 字的產業別分析：

產業別：{industry}
月電費：{monthly_electricity_bill_ntd:,.0f} NTD"""
    
    # 如果有碳排數據，加入 prompt
    if emission_total_tco2e and emission_total_tco2e > 0:
        prompt += f"""
年碳排放總額：{emission_total_tco2e:.2f} tCO₂e"""
    
    prompt += f"""

請分析以下內容：
1. 產業別相關規範要求
2. 市場趨勢
3. 風險分析
4. 判斷耗能等級（高耗能/中耗能/低耗能）
5. 【必須估算年營收】根據耗能等級估算年營收：
   - 高耗能：月電費 × 300倍
   - 中耗能：月電費 × 600倍
   - 低耗能：月電費 × 1200倍
   請根據產業特性和耗能等級，估算並明確標示年營收數值"""
    
    # 如果有碳排數據，要求 LLM 在分析中包含
    if emission_total_tco2e and emission_total_tco2e > 0:
        prompt += f"""
6. 【必須包含】年碳排放總額：{emission_total_tco2e:.2f} tCO₂e（必須在分析中明確提及此具體數據）"""
    
    prompt += f"""

請用繁體中文撰寫，並在最後明確標示：
- 耗能等級：[高耗能/中耗能/低耗能]
- 估算年營收：[數值] NTD（LLM 根據耗能等級估算，不顯示計算過程，只顯示結果）"""
    
    if emission_total_tco2e and emission_total_tco2e > 0:
        prompt += f"""
- 年碳排放總額：{emission_total_tco2e:.2f} tCO₂e（必須明確標示）"""
    
    prompt += f"""

格式範例：
「{industry}產業面臨...規範要求...市場趨勢...風險分析...年碳排放總額 {emission_total_tco2e:.2f} tCO₂e...耗能等級：中耗能。計算年營收：54,000,000 NTD。」"""

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
    
    # 整段 LLM 回應（150字），不做任何萃取
    analysis_text = response.content[0].text if response.content else ""
    
    # 構建結果：整段 150 字直接寫入，不做萃取
    result = {
        "industry": industry,
        "monthly_electricity_bill_ntd": monthly_electricity_bill_ntd,
        "industry_analysis": analysis_text.strip(),  # 整段 150 字，不做萃取
        "session_id": session_id,
        "timestamp": datetime.now().isoformat(),
        "step": "Step 1 - 產業別分析"
    }
    
    # 如果有碳排數據，也記錄（但不萃取150字中的內容）
    if emission_total_tco2e and emission_total_tco2e > 0:
        result["emission_total_tco2e"] = emission_total_tco2e
    
    # 寫入 log
    save_industry_analysis_to_log(result)
    
    return result


def save_industry_analysis_to_log(data: Dict[str, Any]) -> None:
    """
    將產業別分析硬寫入 log 文件
    確保整段 150 字 LLM 回應全部寫入
    """
    import os
    
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
            # 合併數據（產業別分析優先，確保新的 150 字覆蓋舊的）
            existing_data.update(data)
            data = existing_data
        except:
            pass
    
    # 硬寫入文件：確保整段 150 字全部寫入
    try:
        with open(log_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()  # 強制刷新緩衝區
            os.fsync(f.fileno())  # 強制寫入磁盤
        print(f"[產業別分析] 已硬寫入 log: {log_file.name} (整段 {len(data.get('industry_analysis', ''))} 字)")
    except Exception as e:
        raise Exception(f"硬寫入產業別分析 log 失敗: {e}")

