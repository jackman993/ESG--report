"""Step 2: 生成公司環境段報告"""
import streamlit as st
import sys
import json
from pathlib import Path
from datetime import datetime

# 導入共享模組
sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import *
from shared.utils import render_output_folder_links, render_api_key_input

# ============ 後台 Log 函數 ============
def save_session_log(session_data):
    """儲存用戶 session log 到後台"""
    session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = BACKEND_LOGS / f"session_{session_id}.json"
    
    session_data["session_id"] = session_id
    session_data["timestamp"] = datetime.now().isoformat()
    
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(session_data, f, ensure_ascii=False, indent=2)
    
    print(f"  ✓ Session log 已儲存: {log_file.name}")
    return log_file

# 頁面配置
st.set_page_config(page_title="Step 2: 公司環境段報告", page_icon="🌱", layout="wide")

# 側邊欄（使用 Streamlit 自動導航，只保留 API Key 和輸出文件夾）
st.sidebar.divider()
API_KEY = render_api_key_input()
render_output_folder_links()

# 主頁面
st.title("🌱 Step 2: 生成公司環境段報告")

# 前置條件檢查
st.subheader("📋 前置條件檢查")
step1_done = st.session_state.get("step1_done", False)
if step1_done:
    st.success("✅ Step 1 已完成")
else:
    st.warning("⬜ 請先完成 Step 1")

st.divider()

# 環境段生成
if step1_done:
    st.subheader("📑 環境段生成")
    st.info("生成 17 頁環境章節報告")
    
    # 檢查前置條件
    tcfd_done = "results" in st.session_state and st.session_state.results
    emission_done = st.session_state.get("emission_done", False)
    
    if not tcfd_done:
        st.warning("⚠️ 請先完成 Step 1 的 TCFD 表格生成")
    elif not emission_done:
        st.warning("⚠️ 請先完成 Step 1 的碳排計算")
    else:
        # 測試模式選項
        test_mode = st.checkbox("🧪 測試模式（跳過 LLM API，快速預覽）", value=False)
        
        if st.button("🚀 一鍵生成環境章節 (17頁 PPTX)", type="primary", use_container_width=True, key="btn_env"):
            with st.spinner("生成中...請稍候（約 2-3 分鐘）"):
                try:
                    # 加入 environment report 路徑（從 ESG go 目錄）
                    BASE_DIR = Path(__file__).parent.parent.parent  # ESG go/
                    env_report_path = BASE_DIR / "environment report"
                    sys.path.insert(0, str(env_report_path))
                    
                    from environment_pptx import EnvironmentPPTXEngine
                    from datetime import datetime
                    
                    st.info("📄 正在調用 Environment PPTX 引擎...")
                    
                    # 使用統一的 API Key
                    import config
                    config.ANTHROPIC_API_KEY = API_KEY
                    
                    # 取得 Step 1 的 TCFD 資料夾
                    tcfd_output_folder = st.session_state.get("tcfd_output_folder", None)
                    emission_data = st.session_state.get("emission_data", {})
                    # 優先使用 industry_selected，如果沒有則使用 widget 的值
                    industry_name = st.session_state.get("industry_selected") or st.session_state.get("industry", "企業")
                    emission_output_folder = st.session_state.get("emission_output_folder", str(OUTPUT_B_EMISSION))
                    company_profile = st.session_state.get("company_profile", {})
                    
                    # 模板路徑（從 ESG go 目錄）
                    BASE_DIR = Path(__file__).parent.parent.parent  # ESG go/
                    template_path = BASE_DIR / "environment report" / "assets" / "templet_english.pptx"
                    
                    # 生成報告
                    engine = EnvironmentPPTXEngine(
                        template_path=template_path,
                        test_mode=test_mode,
                        emission_data=emission_data,
                        industry=industry_name,
                        tcfd_output_folder=tcfd_output_folder,
                        emission_output_folder=emission_output_folder,
                        company_profile=company_profile,
                        api_key=API_KEY
                    )
                    report = engine.generate()
                    
                    # 儲存到 C_Environment 資料夾
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"ESG環境篇_{timestamp}.pptx"
                    output_path = OUTPUT_C_ENVIRONMENT / output_filename
                    
                    OUTPUT_C_ENVIRONMENT.mkdir(parents=True, exist_ok=True)
                    engine.save(str(output_path))
                    
                    if output_path.exists():
                        st.success(f"✅ 檔案已儲存！")
                        st.info(f"📁 **完整路徑：** `{output_path}`")
                        st.session_state.step2_done = True
                        
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
                            "step": "Step 2",
                            "industry": industry_name,
                            "company_profile": company_profile,
                            "emission_data": emission_data,
                            "tcfd_output_folder": tcfd_output_folder,
                            "output_path": str(output_path),
                            "test_mode": test_mode,
                            "slide_count": len(report.slides)
                        }
                        save_session_log(session_log)
                    else:
                        st.error(f"❌ 檔案儲存失敗！路徑：{output_path}")
                        
                except Exception as e:
                    st.error(f"❌ 生成失敗：{e}")
                    st.exception(e)
else:
    st.info("請先完成 Step 1 後再生成環境段報告")

