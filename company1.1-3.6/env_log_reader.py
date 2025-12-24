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
    讀取最新的環境段 log 檔案並合併同一個 session 的資料
    此函數會：
    1. 找到最新的 log 檔案
    2. 從該檔案中提取 session_id
    3. 找到所有相同 session_id 的 log 檔案
    4. 合併所有檔案的資料（特別是從 Step 1 獲取產業別）
    
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
        # 先讀取最新的檔案
        with open(latest_file, "r", encoding="utf-8") as f:
            latest_data = json.load(f)
        
        # 從最新檔案中提取 session_id
        session_id = latest_data.get("session_id", "")
        
        # 如果我們有 session_id，嘗試從同一個 session 的其他檔案中尋找並合併資料
        merged_data = latest_data.copy()
        
        if session_id:
            # 找到所有相同 session_id 的檔案
            session_files = [f for f in json_files if f != latest_file]
            
            for session_file in session_files:
                try:
                    with open(session_file, "r", encoding="utf-8") as f:
                        session_data = json.load(f)
                    
                    # 檢查此檔案是否屬於同一個 session
                    if session_data.get("session_id") == session_id:
                        # 合併資料：優先從 Step 1 檔案中獲取產業別和 company_profile
                        # 必須檢查值是否為非空
                        if "industry" in session_data:
                            industry_value = session_data["industry"]
                            if industry_value and str(industry_value).strip():
                                if not merged_data.get("industry") or not str(merged_data.get("industry", "")).strip():
                                    merged_data["industry"] = str(industry_value).strip()
                                    print(f"[INFO] Merged industry '{merged_data['industry']}' from: {session_file.name}")
                        
                        if "company_profile" in session_data:
                            if "company_profile" not in merged_data:
                                merged_data["company_profile"] = {}
                            # 合併 company_profile，保留現有值除非缺失
                            for key, value in session_data["company_profile"].items():
                                if key not in merged_data["company_profile"]:
                                    merged_data["company_profile"][key] = value
                        
                        # 合併 emission_data（如果可用）
                        if "emission_data" in session_data:
                            if "emission_data" not in merged_data:
                                merged_data["emission_data"] = {}
                            for key, value in session_data["emission_data"].items():
                                if key not in merged_data["emission_data"]:
                                    merged_data["emission_data"][key] = value
                        
                        # 合併 tcfd_summary（如果可用）
                        if "tcfd_summary" in session_data:
                            if "tcfd_summary" not in merged_data:
                                merged_data["tcfd_summary"] = {}
                            for key, value in session_data["tcfd_summary"].items():
                                if key not in merged_data["tcfd_summary"]:
                                    merged_data["tcfd_summary"][key] = value
                        
                        print(f"[INFO] Merged data from: {session_file.name}")
                
                except Exception as e:
                    # 跳過無法讀取的檔案
                    continue
        
        # 如果仍然缺少關鍵資料，嘗試從最近的包含完整資料的檔案中尋找
        # 檢查我們缺少什麼（必須檢查值是否為非空）
        missing_industry = not merged_data.get("industry") or not str(merged_data.get("industry", "")).strip()
        missing_company_profile = "company_profile" not in merged_data or not merged_data.get("company_profile")
        missing_emission = "emission_data" not in merged_data or not merged_data.get("emission_data")
        missing_tcfd = "tcfd_summary" not in merged_data or not merged_data.get("tcfd_summary")
        
        if missing_industry or missing_company_profile or missing_emission or missing_tcfd:
            # 優先搜尋 Step 1 的文件（產業別通常在 Step 1 中記錄）
            # 先按 step 排序（Step 1 優先），然後按修改時間排序
            def sort_key(f):
                try:
                    with open(f, "r", encoding="utf-8") as file:
                        data = json.load(file)
                        step = str(data.get("step", "")).lower()
                        # Step 1 優先，然後按時間排序
                        if "step 1" in step:
                            return (0, f.stat().st_mtime)
                        else:
                            return (1, f.stat().st_mtime)
                except:
                    return (2, f.stat().st_mtime)
            
            # 搜尋最近的檔案（Step 1 優先，然後按修改時間排序）
            for log_file in sorted(json_files, key=sort_key):
                if log_file == latest_file:
                    continue  # 跳過我們已經讀取的檔案
                
                try:
                    with open(log_file, "r", encoding="utf-8") as f:
                        file_data = json.load(f)
                    
                    # 合併產業別（如果缺失）- 必須檢查值是否為非空
                    if missing_industry and "industry" in file_data:
                        industry_value = file_data["industry"]
                        # 檢查產業別是否為非空（不是空字串、None 或只包含空白）
                        if industry_value and str(industry_value).strip():
                            merged_data["industry"] = str(industry_value).strip()
                            print(f"[INFO] Found industry '{merged_data['industry']}' from: {log_file.name} (step: {file_data.get('step', 'unknown')})")
                            missing_industry = False
                    
                    # 合併 company_profile（如果缺失）
                    if missing_company_profile and "company_profile" in file_data:
                        if "company_profile" not in merged_data:
                            merged_data["company_profile"] = {}
                        for key, value in file_data["company_profile"].items():
                            if key not in merged_data["company_profile"]:
                                merged_data["company_profile"][key] = value
                        if merged_data["company_profile"]:
                            print(f"[INFO] Found company_profile from: {log_file.name}")
                            missing_company_profile = False
                    
                    # 合併 emission_data（如果缺失）
                    if missing_emission and "emission_data" in file_data:
                        if "emission_data" not in merged_data:
                            merged_data["emission_data"] = {}
                        for key, value in file_data["emission_data"].items():
                            if key not in merged_data["emission_data"]:
                                merged_data["emission_data"][key] = value
                        if merged_data["emission_data"]:
                            print(f"[INFO] Found emission_data from: {log_file.name}")
                            missing_emission = False
                    
                    # 合併 tcfd_summary（如果缺失）
                    if missing_tcfd and "tcfd_summary" in file_data:
                        if "tcfd_summary" not in merged_data:
                            merged_data["tcfd_summary"] = {}
                        for key, value in file_data["tcfd_summary"].items():
                            if key not in merged_data["tcfd_summary"]:
                                merged_data["tcfd_summary"][key] = value
                        if merged_data["tcfd_summary"]:
                            print(f"[INFO] Found tcfd_summary from: {log_file.name}")
                            missing_tcfd = False
                    
                    # 如果我們找到了所有缺失的資料，停止搜尋
                    if not (missing_industry or missing_company_profile or missing_emission or missing_tcfd):
                        break
                
                except Exception:
                    continue
        
        # 轉換為標準格式
        standardized = _standardize_log_data(merged_data)
        
        print(f"[OK] Loaded environment log: {latest_file.name}")
        if session_id:
            print(f"[INFO] Session ID: {session_id}")
        if standardized.get("industry"):
            print(f"[INFO] Industry: {standardized['industry']}")
        
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
    
    # 產業別（如果沒有找到或為空，保持空字串而不是設置默認值）
    industry_raw = raw_data.get("industry", "")
    # 確保產業別不是 None 且去除空白後不為空
    if industry_raw and str(industry_raw).strip():
        standardized["industry"] = str(industry_raw).strip()
    else:
        standardized["industry"] = ""
    
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
    
    # 碳排總額（支援 emission_result 和 emission_data 兩種格式）
    emission_result = raw_data.get("emission_result", {})
    emission_data = raw_data.get("emission_data", {})
    
    # 優先嘗試 emission_result，然後嘗試 emission_data
    if emission_result and "total" in emission_result:
        standardized["emission_total_tco2e"] = emission_result.get("total", 0.0)
    elif emission_data and "total" in emission_data:
        standardized["emission_total_tco2e"] = emission_data.get("total", 0.0)
    else:
        standardized["emission_total_tco2e"] = 0.0
    
    # 公司名稱
    standardized["company_name"] = raw_data.get("company_name", "").strip()
    
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
            "company_name": "",
            "company_context": "",
            "tcfd_policy_context": "",
            "tcfd_market_context": "",
            "emission_context": "",
        }
    
    # 建立公司背景上下文
    print(f"[DEBUG] get_prompt_context: 開始提取上下文")
    print(f"[DEBUG] get_prompt_context: log_data 是否為 None: {log_data is None}")
    if log_data:
        print(f"[DEBUG] get_prompt_context: log_data keys: {list(log_data.keys())}")
        print(f"[DEBUG] get_prompt_context: log_data['industry'] 原始值: {repr(log_data.get('industry', 'KEY NOT FOUND'))}")
        print(f"[DEBUG] get_prompt_context: log_data['industry'] 類型: {type(log_data.get('industry'))}")
        print(f"[DEBUG] get_prompt_context: log_data['industry'] 是否為空字串: {log_data.get('industry', '') == ''}")
        print(f"[DEBUG] get_prompt_context: log_data['industry'] 是否為 None: {log_data.get('industry') is None}")
    
    industry = log_data.get("industry", "")
    company_name = log_data.get("company_name", "").strip()
    revenue_display = log_data.get("estimated_revenue_display", "")
    company_size = log_data.get("company_size", "")
    
    # 調試信息：確認產業別是否成功提取
    print(f"[DEBUG] get_prompt_context: 提取後的 industry 變數: {repr(industry)}")
    print(f"[DEBUG] get_prompt_context: industry 類型: {type(industry)}")
    print(f"[DEBUG] get_prompt_context: industry 是否為空字串: {industry == ''}")
    print(f"[DEBUG] get_prompt_context: industry 是否為 None: {industry is None}")
    if not industry:
        print(f"[ERROR] get_prompt_context: 未能從 log_data 中提取產業別！")
        print(f"[DEBUG] log_data keys: {list(log_data.keys()) if log_data else 'log_data is None'}")
        print(f"[DEBUG] log_data['industry']: {repr(log_data.get('industry', 'KEY NOT FOUND') if log_data else 'log_data is None')}")
    else:
        print(f"[OK] get_prompt_context: 成功提取產業別: {industry}")
    
    company_context = f"公司營運於 {industry} 產業。" if industry else "公司營運資訊。"
    if revenue_display:
        company_context += f" 推估年營收：{revenue_display}。"
    if company_size:
        company_context += f" 公司規模：{company_size}。"
    
    return {
        "industry": industry,
        "company_name": company_name if company_name else "本公司",
        "company_context": company_context.strip(),
        "tcfd_policy_context": log_data.get("tcfd_policy_regulation", ""),
        "tcfd_market_context": log_data.get("tcfd_market_trends", ""),
        "emission_context": f"Total annual carbon emissions: {log_data.get('emission_total_tco2e', 0.0):.2f} tCO₂e.",
    }

