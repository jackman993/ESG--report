"""
Environment log reader for ESGoneclick.
Reads standardized environment log files and provides data for LLM prompts.
"""
import json
import os
import re
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime


# ============ Log 標準格式 ============
# 標準 log 檔案格式（JSON）：
# {
#   "industry": "Food Industry",
#   "monthly_electricity_bill_ntd": 125890.0,
#   "estimated_revenue_ntd": 45320400.0,  # monthly_bill * 12 * 40
#   "estimated_revenue_display": "22.66-30.21 million NTD",
#   "company_size": "Small-Medium Enterprise",
#   "tcfd_policy_regulation": "Carbon tax policies expected to be implemented from 2024-2030...",
#   "tcfd_market_trends": "Climate-conscious consumers expected to reach 65-70% by 2026...",
#   "emission_total_tco2e": 179.02,
#   "session_id": "20251212_231752",
#   "timestamp": "2025-12-12T23:17:52.459138"
# }


# 預設 log 路徑
DEFAULT_LOG_DIR = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")


def load_latest_environment_log(log_dir: Optional[Path] = None) -> Optional[Dict[str, Any]]:
    """
    讀取最新的環境段 log 檔案
    
    Args:
        log_dir: Log 資料夾路徑（預設使用 DEFAULT_LOG_DIR）
    
    Returns:
        解析後的 log 資料 dict，如果找不到則返回 None
    """
    if log_dir is None:
        log_dir = DEFAULT_LOG_DIR
    
    log_dir = Path(log_dir)
    
    if not log_dir.exists():
        print(f"[WARN] Log directory not found: {log_dir}")
        return None
    
    # 尋找所有 JSON 檔案
    json_files = list(log_dir.glob("*.json"))
    
    if not json_files:
        print(f"[WARN] No log files found in: {log_dir}")
        return None
    
    # 按修改時間排序，取最新的
    latest_file = max(json_files, key=lambda f: f.stat().st_mtime)
    
    try:
        with open(latest_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # 轉換為標準格式（如果來源格式不同）
        standardized = _standardize_log_data(data)
        
        print(f"[OK] Loaded environment log: {latest_file.name}")
        return standardized
    
    except Exception as e:
        print(f"[ERROR] Failed to load log file {latest_file}: {e}")
        return None


def _standardize_log_data(raw_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    將原始 log 資料轉換為標準格式
    
    Args:
        raw_data: 原始 log 資料
    
    Returns:
        標準化的 log 資料
    """
    standardized = {}
    
    # 產業別
    standardized["industry"] = raw_data.get("industry", "General Industry")
    
    # 月電費和推估營收
    company_profile = raw_data.get("company_profile", {})
    monthly_bill = company_profile.get("monthly_bill_ntd", 0.0)
    standardized["monthly_electricity_bill_ntd"] = monthly_bill
    
    # 推估營收（月電費 × 12個月 × 40倍）
    estimated_revenue = monthly_bill * 12 * 40
    standardized["estimated_revenue_ntd"] = estimated_revenue
    
    # 營收顯示格式（加上數值標註避免混淆）
    revenue_display = company_profile.get("revenue_display", "")
    if not revenue_display:
        # 計算萬元數值
        revenue_wan = estimated_revenue / 10000
        # 格式化：萬元(數值) NTD，例如：1386萬(13,860,000) NTD
        standardized["estimated_revenue_display"] = f"{revenue_wan:.0f}萬({estimated_revenue:,.0f}) NTD"
    else:
        # 如果已有顯示格式，嘗試解析並加上數值標註
        # 處理 "萬元" 格式
        if "萬元" in revenue_display:
            # 提取萬元數值
            match = re.search(r'([\d,]+\.?\d*)\s*萬元', revenue_display)
            if match:
                wan_value = float(match.group(1).replace(',', ''))
                standardized["estimated_revenue_display"] = f"{wan_value:.0f}萬({estimated_revenue:,.0f}) NTD"
            else:
                standardized["estimated_revenue_display"] = revenue_display.replace("萬元", f"萬({estimated_revenue:,.0f}) NTD")
        else:
            # 其他格式，直接加上數值標註
            standardized["estimated_revenue_display"] = f"{revenue_display} ({estimated_revenue:,.0f} NTD)"
    
    # 公司規模
    standardized["company_size"] = company_profile.get("size", "Unknown")
    
    # TCFD 政策與法規
    tcfd_summary = raw_data.get("tcfd_summary", {})
    standardized["tcfd_policy_regulation"] = tcfd_summary.get(
        "transformation_policy", 
        "No policy and regulation data available."
    )
    
    # TCFD 市場趨勢
    standardized["tcfd_market_trends"] = tcfd_summary.get(
        "market_trend",
        "No market trend data available."
    )
    
    # 碳排總額
    emission_result = raw_data.get("emission_result", {})
    standardized["emission_total_tco2e"] = emission_result.get("total", 0.0)
    
    # Session ID 和時間戳
    standardized["session_id"] = raw_data.get("session_id", "")
    standardized["timestamp"] = raw_data.get("timestamp", "")
    
    return standardized


def get_prompt_context(log_data: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
    """
    從 log 資料中提取用於 LLM prompt 的上下文
    
    Args:
        log_data: Log 資料（如果為 None，則自動讀取最新的）
    
    Returns:
        包含 prompt 上下文的 dict
    """
    if log_data is None:
        log_data = load_latest_environment_log()
    
    if log_data is None:
        return {
            "industry": "",
            "company_context": "",
            "tcfd_policy_context": "",
            "tcfd_market_context": "",
            "emission_context": "",
        }
    
    # 建立公司背景上下文
    industry = log_data.get("industry", "")
    revenue_display = log_data.get("estimated_revenue_display", "")
    company_size = log_data.get("company_size", "")
    
    company_context = f"公司營運於 {industry} 產業。"
    if revenue_display:
        company_context += f" 推估年營收：{revenue_display}。"
    if company_size:
        company_context += f" 公司規模：{company_size}。"
    
    return {
        "industry": industry,
        "company_context": company_context.strip(),
        "tcfd_policy_context": log_data.get("tcfd_policy_regulation", ""),
        "tcfd_market_context": log_data.get("tcfd_market_trends", ""),
        "emission_context": f"Total annual carbon emissions: {log_data.get('emission_total_tco2e', 0.0):.2f} tCO₂e.",
    }

