"""
產業別分析生成器
在 Step 1 用戶按「生成 TCFD」時，第一個 LLM 調用生成產業別分析
"""
import anthropic
import json
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

# 不再從 config 導入模型，直接使用與 TCFD 表格相同的模型

# 使用相對路徑（兼容本地和容器環境）
# 統一使用 TCFD generator/logs（與 TCFD 原來的 log 路徑一致）
# 從當前文件位置計算：company1.1-3.6/industry_analysis.py -> TCFD generator/logs
_current_file = Path(__file__)  # company1.1-3.6/industry_analysis.py
_base_dir = _current_file.parent.parent  # ESG--report/
# 統一使用 TCFD generator/logs（與 save_session_log 一致）
LOG_FILE_BASE = _base_dir / "TCFD generator" / "logs"


def generate_industry_analysis(session_id: str, api_key: str = None, model: str = None) -> Dict[str, Any]:
    """
    生成產業別分析（150字）- 寫死絕對路徑，不抽象
    api_key: 從 Streamlit UI 輸入的 API key（優先使用），如果為 None 則從 config 讀取
    model: 模型名稱（優先使用），如果為 None 則使用與 TCFD 表格相同的模型
    """
    # 優先使用傳入的 api_key，否則從 config 讀取（向後兼容）
    if api_key:
        final_api_key = api_key
    else:
        from config_pptx_company import ANTHROPIC_API_KEY
        final_api_key = ANTHROPIC_API_KEY
    
    if not final_api_key:
        raise RuntimeError("API key is not configured.")
    
    client = anthropic.Anthropic(api_key=final_api_key)
    # 優先使用傳入的 model，否則使用與 TCFD 表格相同的模型
    final_model = model if model else "claude-sonnet-4-20250514"
    
    # 相對路徑讀取 log（兼容所有環境）
    log_file = LOG_FILE_BASE / f"session_{session_id}.json"
    if not log_file.exists():
        raise FileNotFoundError(f"找不到 log 文件: {log_file}")
    with open(log_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    # 直接讀取（寫死路徑，不抽象）
    industry = data["industry"]
    monthly_electricity_bill_ntd = data["monthly_bill_ntd"]
    emission_total_tco2e = data.get("emission_data", {}).get("total")
    
    # 構建 prompt（寫死，不抽象）
    emission_text = f"\n年碳排放總額：{emission_total_tco2e:.2f} tCO₂e" if emission_total_tco2e else ""
    emission_item = f"\n6. 【必須包含】年碳排放總額：{emission_total_tco2e:.2f} tCO₂e（必須在分析中明確提及此具體數據）" if emission_total_tco2e else ""
    emission_final = f"\n- 年碳排放總額：{emission_total_tco2e:.2f} tCO₂e（必須明確標示）" if emission_total_tco2e else ""
    emission_example = f"{emission_total_tco2e:.2f} tCO₂e" if emission_total_tco2e else "XX.XX tCO₂e"
    
    prompt = f"""請根據以下資訊，撰寫約 150 字的產業別分析：

產業別：{industry}
月電費：{monthly_electricity_bill_ntd:,.0f} NTD{emission_text}

請分析以下內容：
1. 產業別相關規範要求
2. 市場趨勢
3. 風險分析
4. 判斷耗能等級（高耗能/中耗能/低耗能）
5. 【必須估算年營收】根據耗能等級估算年營收：
   - 高耗能：月電費 × 300倍
   - 中耗能：月電費 × 600倍
   - 低耗能：月電費 × 1200倍
   請根據產業特性和耗能等級，估算並明確標示年營收數值{emission_item}

請用繁體中文撰寫，並在最後明確標示：
- 耗能等級：[高耗能/中耗能/低耗能]
- 估算年營收：[數值] NTD（LLM 根據耗能等級估算，不顯示計算過程，只顯示結果）{emission_final}

格式範例：
「{industry}產業面臨...規範要求...市場趨勢...風險分析...年碳排放總額 {emission_example}...耗能等級：中耗能。計算年營收：54,000,000 NTD。」"""

    # 調用 LLM
    response = client.messages.create(
        model=final_model,
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
    """硬寫入 log（相對路徑，兼容所有環境）"""
    import os
    
    # 確保目錄存在
    LOG_FILE_BASE.mkdir(parents=True, exist_ok=True)
    
    session_id = data["session_id"]
    log_file = LOG_FILE_BASE / f"session_{session_id}_industry_analysis.json"
    
    # 調試：打印路徑信息
    print(f"[save_industry_analysis_to_log] LOG_FILE_BASE: {LOG_FILE_BASE}")
    print(f"[save_industry_analysis_to_log] log_file: {log_file}")
    print(f"[save_industry_analysis_to_log] log_file.exists(): {log_file.exists()}")
    print(f"[save_industry_analysis_to_log] 準備寫入 {len(data.get('industry_analysis', ''))} 字")
    
    try:
        with open(log_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        print(f"[save_industry_analysis_to_log] ✅ 成功寫入: {log_file}")
        print(f"[save_industry_analysis_to_log] 文件大小: {log_file.stat().st_size} bytes")
    except Exception as e:
        print(f"[save_industry_analysis_to_log] ❌ 寫入失敗: {e}")
        raise

