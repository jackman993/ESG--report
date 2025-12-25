"""Step 1: 碳排與TCFD氣候治理"""
import streamlit as st
import anthropic
import sys
import zipfile
import io
import json
from pathlib import Path
from datetime import datetime

# 導入共享模組
sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import *
from shared.utils import render_output_folder_links, render_api_key_input, render_sidebar_navigation, generate_report_summary, switch_page

# 加入 TCFD_Table 路徑（tcfd_* 模組位於此目錄）
tcfd_table_path = Path(__file__).parent.parent / "TCFD_Table"
sys.path.insert(0, str(tcfd_table_path))

# 從 TCFD_Table 目錄導入 tcfd 模組
from tcfd_01_transformation import create_table as create_01
from tcfd_02_market import create_table as create_02
from tcfd_03_physical import create_table as create_03
from tcfd_04_temperature import create_table as create_04
from tcfd_05_resource import create_table as create_05

# ============ 後台 Log 函數 ============
def save_session_log(session_data):
    """儲存用戶 session log 到 TCFD generator/logs 文件夾"""
    # 使用 TCFD generator/logs 文件夾
    log_dir = Path(__file__).parent.parent / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # 使用現有的 session_id 或生成新的
    session_id = session_data.get("session_id", datetime.now().strftime("%Y%m%d_%H%M%S"))
    log_file = log_dir / f"session_{session_id}.json"
    
    # 如果文件已存在，讀取並合併數據
    if log_file.exists():
        try:
            with open(log_file, "r", encoding="utf-8") as f:
                existing_data = json.load(f)
            # 合併數據（新數據覆蓋舊數據）
            existing_data.update(session_data)
            session_data = existing_data
        except:
            pass
    
    session_data["session_id"] = session_id
    session_data["timestamp"] = datetime.now().isoformat()
    
    # 硬寫入 log 文件
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(session_data, f, ensure_ascii=False, indent=2)
        f.flush()
        import os
        os.fsync(f.fileno())
    
    print(f"[TCFD Log] 已保存到: {log_file.name}")
    
    print(f"  ✓ Session log 已儲存: {log_file.name}")
    return log_file

def calculate_company_profile(monthly_bill_ntd, industry_name):
    """根據月電費估算公司規模和節能投資預算"""
    annual_revenue_ntd = monthly_bill_ntd * 360
    annual_revenue_wan = annual_revenue_ntd / 10000
    
    if annual_revenue_ntd > 100_000_000:
        size = "中大型"
    elif annual_revenue_ntd > 50_000_000:
        size = "中型"
    else:
        size = "中小型"
    
    budget_ntd = annual_revenue_ntd * 0.02
    budget_wan = budget_ntd / 10000
    
    revenue_display = f"{annual_revenue_wan:.0f}萬元 ({annual_revenue_ntd:,.0f})"
    
    return {
        "monthly_bill_ntd": monthly_bill_ntd,
        "annual_revenue_ntd": annual_revenue_ntd,
        "annual_revenue_wan": annual_revenue_wan,
        "revenue_display": revenue_display,
        "revenue_for_prompt": f"{annual_revenue_wan:.0f}萬元",
        "size": size,
        "budget_ntd": budget_ntd,
        "budget_wan": budget_wan,
        "budget_display": f"{budget_wan:.1f}萬元",
        "budget_for_prompt": f"{budget_wan:.1f}萬元"
    }

# 專家角色
EXPERT_ROLE = "你是 ESG 的 GRI 和 TCFD 專家。"

# 5 個表格設定
TABLES = [
    {
        "name": "01 轉型風險",
        "create": create_01,
        "prompt": EXPERT_ROLE + """針對「{industry}」進行 TCFD 轉型風險分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：政策與法規風險
第2行：綠色產品與科技風險
只輸出 2 行，不要其他文字。"""
    },
    {
        "name": "02 市場風險",
        "create": create_02,
        "prompt": EXPERT_ROLE + """針對「{industry}」進行 TCFD 市場風險分析，聚焦 2026 年以後趨勢，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：消費者偏好變化風險
第2行：市場需求變化風險
只輸出 2 行，不要其他文字。"""
    },
    {
        "name": "03 實體風險",
        "create": create_03,
        "prompt": EXPERT_ROLE + """針對「{industry}」進行 TCFD 實體風險分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：極端氣候事件風險
第2行：長期氣候變遷風險
只輸出 2 行，不要其他文字。"""
    },
    {
        "name": "04 溫升風險",
        "create": create_04,
        "prompt": EXPERT_ROLE + """針對「{industry}」進行 TCFD 溫升情境風險分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
風險描述|||財務影響|||因應措施
第1行：升溫1.5°C情境風險
第2行：升溫2°C以上情境風險
只輸出 2 行，不要其他文字。"""
    },
    {
        "name": "05 資源效率",
        "create": create_05,
        "prompt": EXPERT_ROLE + """針對「{industry}」進行 TCFD 資源效率機會分析，用繁體中文回答。
本公司年營收約 {revenue}，請以此規模為基準。
建議短期節能投資以營收的 2% 為基準（約 {budget}）。
金額請以萬元為單位，避免使用億元或億美元。
請詳細分析，每個重點 80~120 字，包含具體數據、比例、時程。
輸出 2 行，每行用 ||| 分隔三欄，每欄 3 點用分號(;)隔開：
機會描述|||潛在效益|||行動方案
第1行：能源效率提升機會
第2行：資源循環利用機會
只輸出 2 行，不要其他文字。"""
    },
]

# ============ 頁面配置 ============
st.set_page_config(page_title="Step 1: 碳排與TCFD氣候治理", page_icon="🌍", layout="wide")

# 側邊欄（自定義導航）
render_sidebar_navigation()
st.sidebar.divider()
API_KEY = render_api_key_input()
render_output_folder_links()

# 主頁面
st.title("🌍 Step 1: 碳排與TCFD氣候治理")

# 輸入區域
st.subheader("📝 基本資訊")
col1, col2 = st.columns(2)
with col1:
    industry = st.text_input("🏭 產業名稱", placeholder="例如：鋁建材業", key="industry")
with col2:
    monthly_bill = st.number_input("💰 月電費（NTD）", value=0.0, min_value=0.0, key="monthly_bill")

st.divider()

# 子步驟1: 碳排計算
st.subheader("🌱 子步驟1: 碳排放計算")
st.info("使用月電費計算碳排放量")

# 導入 Emission 引擎（從 ESG go 目錄的 emission 資料夾）
BASE_DIR = Path(__file__).parent.parent.parent  # ESG--report/
EMISSION_ENGINE_PATH = BASE_DIR / "emission"
sys.path.insert(0, str(EMISSION_ENGINE_PATH))
from emission_calc import Inputs, estimate

# 選擇模式
calc_mode = st.radio("估算模式", ["Quick (80%)", "Detail (95%)"], horizontal=True, key="calc_mode")

if "Quick" in calc_mode:
    st.markdown("#### Quick 模式：只需月電費")
    default_bill = monthly_bill if monthly_bill > 0 else 50000.0
    emission_monthly_bill = st.number_input("月電費（NTD）", value=default_bill, key="quick_bill")
else:
    st.markdown("#### Detail 模式：完整輸入")
    col1, col2 = st.columns(2)
    with col1:
        default_bill = monthly_bill if monthly_bill > 0 else 50000.0
        emission_monthly_bill = st.number_input("月電費（NTD）", value=default_bill, key="detail_bill")
        price_per_kwh = st.number_input("每度電價（NTD）", value=4.4, key="price_kwh")
        annual_kwh = st.number_input("年用電量（kWh，選填）", value=0.0, key="annual_kwh")
    with col2:
        car_count = st.number_input("汽車台數", value=2, key="car_count")
        motorcycles = st.number_input("機車台數", value=5, key="mc_count")
        gas_liters = st.number_input("汽油（L/年）", value=0.0, key="gas_l")
        refrigerant_kg = st.number_input("冷媒逸散（kg/年）", value=2.0, key="ref_kg")
        refrigerant_gwp = st.number_input("冷媒 GWP", value=1000.0, key="ref_gwp")

if st.button("🧮 計算碳排放", type="primary", use_container_width=True, key="btn_emission"):
    if not API_KEY:
        st.error("請先在左側輸入 API Key")
        st.stop()
    
    with st.spinner("計算中..."):
        # 建立輸入物件
        if "Quick" in calc_mode:
            inp = Inputs(
                mode="quick",
                monthly_bill_ntd=emission_monthly_bill or None,
                price_per_kwh_ntd=4.4,
                use_rule_of_thumb=True,
            )
        else:
            inp = Inputs(
                mode="detail",
                monthly_bill_ntd=emission_monthly_bill or None,
                price_per_kwh_ntd=price_per_kwh,
                annual_kwh=annual_kwh or None,
                car_count=car_count,
                motorcycles=motorcycles,
                gasoline_liters_year=gas_liters or None,
                refrigerant_leak_kg=refrigerant_kg,
                refrigerant_gwp=refrigerant_gwp,
            )
        
        # 呼叫真正引擎
        result = estimate(inp)
        
        st.success(f"✅ 碳排放計算完成！")
        
        # 顯示結果
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("範疇一", f"{result['Scope1_合計']:.2f} t")
        col2.metric("範疇二", f"{result['Scope2_電力']:.2f} t")
        col3.metric("範疇三", f"{result['Scope3_小項']:.2f} t")
        col4.metric("總排放量", f"{result['總排放_S1S2']:.2f} tCO₂e")
        
        st.markdown(f"**占比：** 電力 {result['占比(%)']['電力']}% | 車輛 {result['占比(%)']['車輛']}% | 冷媒 {result['占比(%)']['冷媒']}%")
        
        # 準備 emission_data_for_pptx（無論 PPTX 生成是否成功都需要）
        emission_data_for_pptx = {
            "scope1": result['Scope1_合計'],
            "scope2": result['Scope2_電力'],
            "total": result['總排放_S1S2'],
            "gasoline": result['Scope1_車輛'],
            "refrigerant": result['Scope1_冷媒'],
            "electricity": result['Scope2_電力'],
            "占比": result['占比(%)']
        }
        
        # 生成 PPTX 表格和圓餅圖
        try:
            BASE_DIR = Path(__file__).parent.parent.parent  # ESG go/
            env_assets_path = BASE_DIR / "environment report" / "assets"
            sys.path.insert(0, str(env_assets_path))
            
            from emission_pptx import set_emission_data, create_emission_table_pptx, create_emission_pie_chart
            
            set_emission_data(emission_data_for_pptx)
            
            # 生成表格 PPTX
            table_path = OUTPUT_B_EMISSION / f"Emission_Table_{result['總排放_S1S2']:.0f}t.pptx"
            create_emission_table_pptx(str(table_path))
            st.success(f"✅ 表格已生成：{table_path.name}")
            
            # 生成圓餅圖
            pie_path = OUTPUT_B_EMISSION / "Emission_PieChart.png"
            create_emission_pie_chart(str(pie_path))
            st.success(f"✅ 圓餅圖已生成：{pie_path.name}")
            
        except Exception as e:
            st.warning(f"⚠️ PPTX 生成略過：{e}")
        
        # 計算公司規模
        industry_for_calc = industry if industry else "企業"
        company_profile = calculate_company_profile(emission_monthly_bill, industry_for_calc)
        
        st.success(f"📊 企業規模：{company_profile['size']}（年營收約 {company_profile['revenue_display']}）")
        st.info(f"💰 建議節能投資預算：{company_profile['budget_display']}")
        
        # 儲存到 session_state
        st.session_state.emission_done = True
        st.session_state.emission_data = emission_data_for_pptx
        st.session_state.emission_output_folder = str(OUTPUT_B_EMISSION)
        st.session_state.company_profile = company_profile
        # 注意：industry 已經由 widget 自動管理，不需要手動設置
        # 使用 industry_selected 來儲存計算時使用的產業名稱
        if industry_for_calc:
            st.session_state.industry_selected = industry_for_calc
        st.session_state.monthly_bill_from_step1 = emission_monthly_bill
        
        # 保存 session log（產業、月電費、碳排數據）到 TCFD generator/logs
        session_id = st.session_state.get("session_id", datetime.now().strftime("%Y%m%d_%H%M%S"))
        st.session_state.session_id = session_id
        
        session_log = {
            "session_id": session_id,
            "step": "Step 1 - 子步驟1",
            "industry": industry_for_calc,
            "monthly_bill": emission_monthly_bill,
            "monthly_bill_ntd": emission_monthly_bill,  # 明確標示
            "company_profile": company_profile,
            "emission_data": emission_data_for_pptx,
            "emission_result": {  # 備用路徑
                "total": emission_data_for_pptx.get("total", 0.0)
            },
            "calc_mode": calc_mode
        }
        save_session_log(session_log)
        st.info(f"📝 已保存 log 到 TCFD generator/logs/session_{session_id}.json")

st.divider()

# 子步驟2: TCFD 表格生成
st.subheader("📊 子步驟2: TCFD 表格生成")
st.info("生成 5 個 TCFD 氣候風險表格")

if st.button("🚀 生成 5 個 TCFD 表格", type="primary", use_container_width=True, key="btn_tcfd"):
    # 驗證 API Key
    if not API_KEY or not API_KEY.strip():
        st.error("❌ 請先在左側輸入有效的 API Key")
        st.stop()
    
    # 驗證 API Key 格式（Anthropic API key 通常以 sk-ant- 開頭）
    if not API_KEY.startswith("sk-ant-"):
        st.warning("⚠️ API Key 格式可能不正確（應以 sk-ant- 開頭）")
        # 不停止，讓用戶嘗試
    
    if not industry:
        st.error("請輸入產業")
        st.stop()
    
    if not monthly_bill or monthly_bill <= 0:
        st.error("請輸入月電費")
        st.stop()
    
    # 檢查是否已完成子步驟1（碳排計算）
    emission_data = st.session_state.get("emission_data", {})
    if not emission_data:
        st.error("❌ 請先完成子步驟1的碳排計算")
        st.stop()
    
    # 取得 session_id（從 session_state 或生成新的）
    session_id = st.session_state.get("session_id", datetime.now().strftime("%Y%m%d_%H%M%S"))
    st.session_state.session_id = session_id
    
    # 計算公司規模
    company_profile = calculate_company_profile(monthly_bill, industry)
    st.info(f"📊 企業規模：{company_profile['size']}（年營收約 {company_profile['revenue_display']}）")
    
    # 儲存到 session_state
    st.session_state.monthly_bill_from_step1 = monthly_bill
    st.session_state.company_profile = company_profile
    # 注意：industry 已經由 widget 自動管理，不需要手動設置
    # 使用 industry_selected 來儲存（如果需要的話）
    if industry:
        st.session_state.industry_selected = industry
    
    # 確保 log 已保存（更新 log 包含最新數據）
    session_log_update = {
        "session_id": session_id,
        "step": "Step 1 - 子步驟2 (TCFD表格生成前)",
        "industry": industry,
        "monthly_bill": monthly_bill,
        "monthly_bill_ntd": monthly_bill,
        "company_profile": company_profile,
        "emission_data": emission_data
    }
    save_session_log(session_log_update)
    
    # ========== 開始生成 TCFD 表格 ==========
    
    # 初始化 Anthropic client（加入錯誤處理）
    try:
        client = anthropic.Anthropic(api_key=API_KEY.strip())
    except Exception as e:
        st.error(f"❌ API Key 初始化失敗：{str(e)}")
        st.info("💡 請檢查 API Key 是否正確，或前往 https://console.anthropic.com/ 獲取新的 API Key")
        st.stop()
    
    results = []
    tcfd_summary = {}
    
    progress_bar = st.progress(0)
    
    # 準備 prompt 參數
    prompt_params = {
        "industry": industry,
        "revenue": company_profile["revenue_for_prompt"],
        "budget": company_profile["budget_for_prompt"]
    }
    
    for idx, table in enumerate(TABLES):
        st.info(f"⏳ {table['name']}...")
        
        # LLM（加入錯誤處理）
        try:
            response = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1024,
                messages=[{"role": "user", "content": table["prompt"].format(**prompt_params)}]
            )
            llm_output = response.content[0].text.strip()
            lines = [line.strip() for line in llm_output.split('\n') if line.strip() and '|||' in line]
            
            # 偵錯：如果沒有解析到資料
            if len(lines) == 0:
                st.warning(f"⚠️ {table['name']} LLM 回傳格式異常，重試中...")
                with st.expander(f"LLM 原始回應 - {table['name']}"):
                    st.code(llm_output)
                # 重試一次
                try:
                    response = client.messages.create(
                        model="claude-sonnet-4-20250514",
                        max_tokens=1024,
                        messages=[{"role": "user", "content": table["prompt"].format(**prompt_params)}]
                    )
                    llm_output = response.content[0].text.strip()
                    lines = [line.strip() for line in llm_output.split('\n') if line.strip() and '|||' in line]
                except anthropic.AuthenticationError as auth_err:
                    st.error(f"❌ API 認證失敗：{str(auth_err)}")
                    st.info("💡 請檢查 API Key 是否正確或已過期")
                    st.stop()
                except Exception as retry_err:
                    st.error(f"❌ 重試失敗：{str(retry_err)}")
                    st.stop()
        except anthropic.AuthenticationError as auth_err:
            st.error(f"❌ API 認證失敗：{str(auth_err)}")
            st.info("💡 請檢查 API Key 是否正確或已過期。前往 https://console.anthropic.com/ 確認 API Key 狀態")
            st.stop()
        except anthropic.APIError as api_err:
            st.error(f"❌ API 調用失敗：{str(api_err)}")
            st.info("💡 可能是 API 配額用盡或服務暫時不可用，請稍後再試")
            st.stop()
        except Exception as e:
            st.error(f"❌ 發生錯誤：{str(e)}")
            st.stop()
        
        # 擷取 TCFD 摘要
        if idx == 0 and lines:  # 01 轉型風險
            first_line = lines[0].split("|||")
            if len(first_line) >= 1:
                policy_desc = first_line[0].strip()[:200]
                tcfd_summary["transformation_policy"] = policy_desc
                tcfd_summary["transformation_raw"] = llm_output
        
        if idx == 1 and lines:  # 02 市場風險
            first_line = lines[0].split("|||")
            if len(first_line) >= 1:
                market_desc = first_line[0].strip()[:200]
                tcfd_summary["market_trend"] = market_desc
                tcfd_summary["market_raw"] = llm_output
        
        # 生成 PPTX
        filepath = table["create"](lines, industry, output_dir=OUTPUT_A_TCFD)
        
        # 讀取檔案內容
        with open(filepath, "rb") as f:
            file_data = f.read()
        
        results.append({
            "name": table["name"], 
            "path": filepath,
            "filename": filepath.name,
            "data": file_data
        })
        st.success(f"✅ {table['name']} 完成（{len(lines)} 行資料）")
        
        progress_bar.progress((idx + 1) / len(TABLES))
    
    # 儲存結果到 session_state
    st.session_state.results = results
    st.session_state.tcfd_summary = tcfd_summary
    
    # 顯示擷取的摘要
    if tcfd_summary:
        with st.expander("📋 TCFD 摘要（供後續 LLM 使用）"):
            if "transformation_policy" in tcfd_summary:
                st.markdown(f"**轉型風險/法規政策：** {tcfd_summary['transformation_policy'][:100]}...")
            if "market_trend" in tcfd_summary:
                st.markdown(f"**市場風險/趨勢：** {tcfd_summary['market_trend'][:100]}...")
    
    # 儲存 TCFD 輸出資料夾路徑
    if results:
        tcfd_output_folder = str(results[0]["path"].parent)
        st.session_state.tcfd_output_folder = tcfd_output_folder
        st.info(f"📁 TCFD 輸出資料夾：{tcfd_output_folder}")
    
    st.balloons()
    st.session_state.step1_done = True
    
    # ========== 【+1 步驟 - 王子路徑：TCFD 5 個表格完成後，第 6 個步驟（只 log，不輸出 pptx）】==========
    st.info("👑 【王子路徑】正在生成產業別分析（150字）- TCFD 5 個表格完成後的第 6 個步驟...")
    
    # 取得 session_id
    session_id = st.session_state.get("session_id", datetime.now().strftime("%Y%m%d_%H%M%S"))
    st.session_state.session_id = session_id
    
    try:
        # 導入 industry_analysis 模組
        current_file = Path(__file__)  # TCFD generator/pages/1_🌍_碳排與TCFD氣候治理.py
        base_dir = current_file.parent.parent  # TCFD generator -> ESG--report
        
        # 嘗試多種路徑計算方式（兼容本地和 Streamlit Cloud）
        possible_paths = [
            base_dir / "company1.1-3.6",  # 從 TCFD generator 向上到根目錄
            Path.cwd() / "company1.1-3.6",  # 從當前工作目錄
            current_file.parent.parent.parent / "company1.1-3.6",  # 如果 base_dir 計算錯誤
        ]
        
        company_path = None
        for path in possible_paths:
            if path.exists() and (path / "industry_analysis.py").exists():
                company_path = path
                break
        
        if company_path is None:
            raise ImportError(f"找不到 company1.1-3.6 目錄。嘗試的路徑: {[str(p) for p in possible_paths]}")
        
        # 清除緩存
        if 'industry_analysis' in sys.modules:
            del sys.modules['industry_analysis']
        
        # 最簡單的導入方式
        if str(company_path) not in sys.path:
            sys.path.insert(0, str(company_path))
        from industry_analysis import generate_industry_analysis, LOG_FILE_BASE
        
        # 調用函數（傳入 session_id、API_KEY 和 model）- 只寫入 log，不生成 pptx
        # 使用 Streamlit UI 輸入的 API_KEY 和與 TCFD 表格相同的模型
        industry_analysis_data = generate_industry_analysis(
            session_id=session_id, 
            api_key=API_KEY.strip(),
            model="claude-sonnet-4-20250514"  # 與 TCFD 5 個表格使用相同的模型
        )
        
        analysis_text = industry_analysis_data.get("industry_analysis", "")
        analysis_length = len(analysis_text) if analysis_text else 0
        
        # 檢查 log 文件是否存在（直接使用 industry_analysis.py 中的 LOG_FILE_BASE，不重新計算）
        log_file = LOG_FILE_BASE / f"session_{session_id}_industry_analysis.json"
        
        if log_file.exists():
            st.success(f"✅ 【王子路徑】產業別分析已生成並寫入 log（{analysis_length}字）- 這是第 6 個步驟（只 log，不輸出 pptx）")
            
            # 顯示 log 文件信息，讓用戶可以確認並閱讀
            with st.expander("📄 查看 150 字產業別分析（Log 文件內容）", expanded=True):
                st.write(f"**📁 Log 文件位置：** `{log_file}`")
                
                # 讀取並顯示完整 log 文件內容
                try:
                    with open(log_file, "r", encoding="utf-8") as f:
                        log_data = json.load(f)
                    
                    st.write("**📊 Log 文件內容：**")
                    st.json(log_data)
                    
                    # 重點顯示 150 字分析
                    if "industry_analysis" in log_data:
                        st.divider()
                        st.write("**📝 150 字產業別分析（核心內容）：**")
                        st.text_area(
                            "產業別分析內容",
                            value=log_data["industry_analysis"],
                            height=200,
                            disabled=True,
                            label_visibility="collapsed"
                        )
                        st.write(f"**字數：** {len(log_data['industry_analysis'])} 字")
                    
                    # 顯示其他關鍵信息
                    if "industry" in log_data:
                        st.write(f"**產業別：** {log_data['industry']}")
                    if "monthly_electricity_bill_ntd" in log_data:
                        st.write(f"**月電費：** {log_data['monthly_electricity_bill_ntd']:,.0f} NTD")
                    if "emission_total_tco2e" in log_data:
                        st.write(f"**年碳排放總額：** {log_data['emission_total_tco2e']:.2f} tCO₂e")
                    if "timestamp" in log_data:
                        st.write(f"**生成時間：** {log_data['timestamp']}")
                        
                except Exception as e:
                    st.error(f"❌ 讀取 log 文件失敗: {e}")
        else:
            st.error(f"❌ 【王子路徑】Log 文件未找到: {log_file}")
            st.write(f"📁 檢查目錄: {log_dir} (存在: {log_dir.exists()})")
            st.write(f"💡 提示：請檢查 industry_analysis.py 中的 LOG_FILE_BASE 路徑計算")
    except Exception as e:
        # 王子路徑失敗不停止流程，只記錄錯誤
        st.error(f"❌ 【王子路徑】產業別分析生成失敗（不影響 TCFD 表格）: {e}")
        st.exception(e)  # 顯示完整錯誤堆棧
        # 不調用 st.stop()，讓流程繼續
    
    # 保存 session log
    session_log = {
        "step": "Step 1 - 子步驟2",
        "industry": industry,
        "company_profile": company_profile,
        "tcfd_summary": tcfd_summary,
        "tcfd_output_folder": tcfd_output_folder if results else None,
        "results_count": len(results)
    }
    save_session_log(session_log)
    
    # 下載區
    st.subheader("📁 下載 TCFD 報告")
    
    # 打包全部下載 (ZIP)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for r in results:
            zip_file.writestr(r["filename"], r["data"])
    zip_buffer.seek(0)
    
    st.download_button(
        label="📦 一次下載全部 (ZIP)",
        data=zip_buffer.getvalue(),
        file_name=f"TCFD_{industry}_全部報告.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary"
    )
    
    st.divider()
    st.write("或個別下載：")
    
    # 個別下載
    cols = st.columns(2)
    for idx, r in enumerate(results):
        with cols[idx % 2]:
            st.download_button(
                label=f"⬇️ {r['name']}", 
                data=r["data"], 
                file_name=r["filename"], 
                key=f"download_{idx}",
                use_container_width=True
            )

st.divider()

# 子步驟3: 生成環境治理報告（第四章）
st.subheader("📑 子步驟3: 生成環境治理報告（第四章）")
st.info("生成 17 頁環境章節報告")

# 檢查前置條件
tcfd_done = "results" in st.session_state and st.session_state.results
emission_done = st.session_state.get("emission_done", False)

if not tcfd_done:
    st.warning("⚠️ 請先完成子步驟2的 TCFD 表格生成")
elif not emission_done:
    st.warning("⚠️ 請先完成子步驟1的碳排計算")
else:
    # 測試模式選項
    test_mode = st.checkbox("🧪 測試模式（跳過 LLM API，快速預覽）", value=False, key="test_mode_step1")
    
    # 檢查是否已經生成過（持久化顯示）
    if "step1_output_path" in st.session_state:
        output_path = Path(st.session_state.step1_output_path)
        if output_path.exists():
            st.markdown("### ✅ 已生成的報告")
            
            # 顯示摘要
            if "step1_summary" in st.session_state:
                st.markdown("### 📝 報告摘要")
                st.info(st.session_state.step1_summary)
            
            st.info(f"📁 **完整路徑：** `{output_path}`")
            
            # 下載按鈕
            with open(output_path, "rb") as f:
                file_data = f.read()
            
            st.download_button(
                label="📥 下載 ESG 環境篇 PPTX",
                data=file_data,
                file_name=st.session_state.get("step1_output_filename", output_path.name),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True
            )
            
            st.success(f"✅ 環境章節生成完成！")
            
            # 下一步按鈕
            st.divider()
            st.markdown("### 🎯 下一步")
            if st.button("➡️ 下一步：重大議題與公司段", use_container_width=True, type="primary", key="next_to_step2_persist"):
                switch_page("pages/4_📋_重大議題段報告.py")
            
            st.divider()
    
    # 生成按鈕（總是顯示，可以重新生成）
    if st.button("🚀 一鍵生成環境章節 (17頁 PPTX)", type="primary", use_container_width=True, key="btn_env_step1"):
        if not API_KEY:
            st.error("請先在左側輸入 API Key")
            st.stop()
        
        with st.spinner("生成中...請稍候（約 2-3 分鐘）"):
            try:
                # 加入 environment report 路徑（從 ESG go 目錄）
                BASE_DIR = Path(__file__).parent.parent.parent  # ESG go/
                env_report_path = BASE_DIR / "environment report"
                sys.path.insert(0, str(env_report_path))
                
                from environment_pptx import EnvironmentPPTXEngine
                from datetime import datetime
                
                st.info("📄 正在調用 Environment PPTX 引擎...")
                
                # 使用統一的 API Key（設置到 environment report 的 config）
                import sys
                env_config_path = BASE_DIR / "environment report" / "config.py"
                if env_config_path.exists():
                    # 動態修改 config 中的 API key
                    import importlib.util
                    spec = importlib.util.spec_from_file_location("config", env_config_path)
                    config = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(config)
                    config.ANTHROPIC_API_KEY = API_KEY
                    # 重新導入以應用更改
                    sys.modules['config'] = config
                
                # 取得 Step 1 的 TCFD 資料夾
                tcfd_output_folder = st.session_state.get("tcfd_output_folder", None)
                emission_data = st.session_state.get("emission_data", {})
                # 優先使用 industry_selected，如果沒有則使用 widget 的值
                industry_name = st.session_state.get("industry_selected") or st.session_state.get("industry", "企業")
                emission_output_folder = st.session_state.get("emission_output_folder", str(OUTPUT_B_EMISSION))
                company_profile = st.session_state.get("company_profile", {})
                
                # 模板路徑（從 ESG go 目錄）
                template_path = BASE_DIR / "environment report" / "assets" / "templet_english.pptx"
                
                # 設置 TCFD 和 Emission 路徑（臨時修改 environment_pptx 的全局變數）
                import environment_pptx as env_pptx_module
                # 更新 TCFD 輸出路徑
                if tcfd_output_folder:
                    env_pptx_module.TCFD_OUTPUT_PATH = tcfd_output_folder
                # 更新 Emission 輸出路徑
                if emission_output_folder:
                    env_pptx_module.EMISSION_OUTPUT_PATH = Path(emission_output_folder)
                
                # 生成報告（傳入 template_path、test_mode 和 api_key）
                engine = EnvironmentPPTXEngine(
                    template_path=str(template_path) if template_path.exists() else None,
                    test_mode=test_mode,
                    api_key=API_KEY,
                    industry=industry_name,
                    company_profile=company_profile,
                    emission_data=emission_data
                )
                report = engine.generate()
                
                # 儲存到 C_Environment 資料夾
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"ESG環境篇_{timestamp}.pptx"
                output_path = OUTPUT_C_ENVIRONMENT / output_filename
                
                OUTPUT_C_ENVIRONMENT.mkdir(parents=True, exist_ok=True)
                engine.save(str(output_path))
                
                if output_path.exists():
                    # 生成摘要（此時會使用 log 中的 150 字分析，已在 TCFD 表格完成後生成）
                    session_id = st.session_state.get("session_id", datetime.now().strftime("%Y%m%d_%H%M%S"))
                    context_data = {
                        "industry": industry_name,
                        "company_profile": company_profile,
                        "emission_data": emission_data,
                        "tcfd_summary": st.session_state.get("tcfd_summary", {}),
                        "session_id": session_id
                    }
                    summary = generate_report_summary("Step 1", context_data, API_KEY, test_mode)
                    
                    # 保存到 session_state（持久化）
                    st.session_state.step1_output_path = str(output_path)
                    st.session_state.step1_summary = summary
                    st.session_state.step1_output_filename = output_filename
                    
                    st.success(f"✅ 檔案已儲存！")
                    st.info(f"📁 **完整路徑：** `{output_path}`")
                    
                    # 顯示摘要
                    st.markdown("### 📝 報告摘要")
                    st.info(summary)
                    
                    # 下載按鈕
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                    
                    st.download_button(
                        label="📥 下載 ESG 環境篇 PPTX",
                        data=file_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True
                    )
                    
                    st.success(f"✅ 環境章節生成完成！共 {len(report.slides)} 頁")
                    st.balloons()
                    
                    # 保存 session log
                    session_log = {
                        "step": "Step 1 - 子步驟3",
                        "industry": industry_name,
                        "company_profile": company_profile,
                        "emission_data": emission_data,
                        "tcfd_output_folder": tcfd_output_folder,
                        "output_path": str(output_path),
                        "summary": summary,
                        "test_mode": test_mode
                    }
                    save_session_log(session_log)
                    
                    # 标记 Step 1 完成
                    st.session_state.step1_done = True
                    
                    # 下一步按鈕
                    st.divider()
                    st.markdown("### 🎯 下一步")
                    if st.button("➡️ 下一步：重大議題與公司段", use_container_width=True, type="primary", key="next_to_step2"):
                        switch_page("pages/4_📋_重大議題段報告.py")
                else:
                    st.error(f"❌ 檔案儲存失敗！路徑：{output_path}")
                    
            except Exception as e:
                st.error(f"❌ 生成失敗：{e}")
                st.exception(e)
