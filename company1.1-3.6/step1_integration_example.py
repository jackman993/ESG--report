"""
Step 1 整合範例
在用戶按「生成 TCFD」按鈕時，先調用產業別分析，然後才生成 TCFD tables
"""

from industry_analysis import generate_industry_analysis
from datetime import datetime

def handle_generate_tcfd_button_click(industry: str, monthly_electricity_bill_ntd: float, session_id: str):
    """
    處理「生成 TCFD」按鈕點擊
    
    流程：
    1. 第一個 LLM 調用：生成產業別分析
    2. 寫入 log
    3. 第二個 LLM 調用：生成 TCFD tables（使用 log 中的產業別分析）
    """
    
    # ========== 第一個 LLM 調用：生成產業別分析 ==========
    print("[Step 1] 開始生成產業別分析...")
    
    industry_analysis_data = generate_industry_analysis(
        industry=industry,
        monthly_electricity_bill_ntd=monthly_electricity_bill_ntd,
        session_id=session_id
    )
    
    print(f"[Step 1] 產業別分析完成：")
    print(f"  - 產業別：{industry_analysis_data['industry']}")
    print(f"  - 耗能等級：{industry_analysis_data['energy_level']}")
    print(f"  - 年營收：{industry_analysis_data['estimated_annual_revenue_display']}")
    print(f"  - 分析內容：{industry_analysis_data['industry_analysis'][:100]}...")
    
    # 數據已自動寫入 log 文件
    # log 文件路徑：C:\Users\User\Desktop\ESG_Output\_Backend\user_logs\session_{session_id}_industry_analysis.json
    
    # ========== 第二個 LLM 調用：生成 TCFD tables ==========
    print("[Step 1] 開始生成 TCFD tables...")
    
    # 這裡調用原有的 TCFD tables 生成函數
    # 該函數會自動從 log 中讀取產業別分析數據
    # generate_tcfd_tables(session_id)  # 原有的函數
    
    print("[Step 1] TCFD tables 生成完成")
    
    return industry_analysis_data


# ========== 使用範例 ==========
if __name__ == "__main__":
    # 模擬用戶輸入
    industry = "焊接業"
    monthly_electricity_bill_ntd = 90000.0
    session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 用戶按「生成 TCFD」按鈕
    result = handle_generate_tcfd_button_click(
        industry=industry,
        monthly_electricity_bill_ntd=monthly_electricity_bill_ntd,
        session_id=session_id
    )
    
    print("\n結果：")
    print(result)

