"""Step 2: 生成重大議題與公司段報告"""
import streamlit as st
import sys
import json
from pathlib import Path
from datetime import datetime

# 導入共享模組
sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import *
from shared.utils import render_output_folder_links, render_api_key_input, render_sidebar_navigation, generate_report_summary, switch_page

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
st.set_page_config(page_title="Step 2: 重大議題與公司段報告", page_icon="📋", layout="wide")

# 側邊欄（自定義導航）
render_sidebar_navigation()
st.sidebar.divider()
API_KEY = render_api_key_input()
render_output_folder_links()

# 主頁面
st.title("📋 Step 2: 生成重大議題與公司段報告")

# 前置條件檢查
st.subheader("📋 前置條件檢查")
step1_done = st.session_state.get("step1_done", False)
if step1_done:
    st.success("✅ Step 1 已完成")
else:
    st.warning("⬜ 請先完成 Step 1")

st.divider()

# 公司段生成
if step1_done:
    st.subheader("🏢 公司段生成")
    
    # 檢查是否已經生成過（持久化顯示）
    if "step2_output_path" in st.session_state:
        output_path = Path(st.session_state.step2_output_path)
        if output_path.exists():
            st.markdown("### ✅ 已生成的報告")
            
            # 顯示摘要
            if "step2_summary" in st.session_state:
                st.markdown("### 📝 報告摘要")
                st.info(st.session_state.step2_summary)
            
            st.info(f"📁 **完整路徑：** `{output_path}`")
            
            # 下載按鈕
            with open(output_path, "rb") as f:
                file_data = f.read()
            
            st.download_button(
                label="📥 下載公司段 PPTX",
                data=file_data,
                file_name=st.session_state.get("step2_output_filename", output_path.name),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True
            )
            
            st.success(f"✅ 公司段生成完成！")
            
            # 下一步按鈕
            st.divider()
            st.markdown("### 🎯 下一步")
            if st.button("➡️ 下一步：治理與社會段", use_container_width=True, type="primary", key="next_to_step3_persist"):
                switch_page("pages/5_🏛️_治理與社會報告.py")
    
    # 公司名稱輸入（可選）
    company_name = st.text_input("公司名稱（可選）", "", key="company_name", 
                                  placeholder="留空則使用「本公司」")
    
    if st.button("🚀 生成公司段 PPTX", type="primary", use_container_width=True, key="btn_company"):
        if not API_KEY:
            st.error("請先在左側輸入 API Key")
            st.stop()
        
        # 使用 progress bar 和 status 來顯示進度，避免卡頓
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.text("📄 正在調用公司段引擎...")
            progress_bar.progress(10)
            
            # 導入包裝器
            from company_engine_wrapper_zh import generate_company_section_zh
            
            progress_bar.progress(30)
            status_text.text("🔄 正在生成公司段內容...")
            
            # 調用包裝器
            output_path, error = generate_company_section_zh(
                api_key=API_KEY,
                company_name=company_name if company_name else None,
                output_dir=OUTPUT_D_COMPANY
            )
            
            progress_bar.progress(90)
            
            if error:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ {error}")
                st.exception(Exception(error))
            else:
                progress_bar.progress(100)
                status_text.empty()
                
                # 生成摘要
                context_data = {
                    "company_name": company_name if company_name else "本公司"
                }
                summary = generate_report_summary("Step 2", context_data, API_KEY, False)
                
                # 保存到 session_state（持久化）
                st.session_state.step2_output_path = str(output_path)
                st.session_state.step2_summary = summary
                st.session_state.step2_output_filename = Path(output_path).name
                
                st.success(f"✅ 公司段生成完成！")
                st.info(f"📁 **完整路徑：** `{output_path}`")
                
                # 顯示摘要
                st.markdown("### 📝 報告摘要")
                st.info(summary)
                
                # 下載按鈕
                if Path(output_path).exists():
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                    
                    st.download_button(
                        label="📥 下載公司段 PPTX",
                        data=file_data,
                        file_name=Path(output_path).name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True
                    )
                    st.balloons()
                    
                    # 保存 session log（包含產業別和 TCFD 市場摘錄）
                    # 從 session_state 取得產業別和 TCFD 摘要
                    industry_name = st.session_state.get("industry_selected") or st.session_state.get("industry", "")
                    tcfd_summary = st.session_state.get("tcfd_summary", {})
                    market_trend = tcfd_summary.get("market_trend", "") if tcfd_summary else ""
                    
                    session_log = {
                        "step": "Step 2",
                        "company_name": company_name if company_name else "本公司",
                        "industry": industry_name,
                        "tcfd_market_trend": market_trend,
                        "output_path": str(output_path),
                        "summary": summary
                    }
                    save_session_log(session_log)
                    
                    # 下一步按鈕
                    st.divider()
                    st.markdown("### 🎯 下一步")
                    if st.button("➡️ 下一步：治理與社會段", use_container_width=True, type="primary", key="next_to_step3"):
                        switch_page("pages/5_🏛️_治理與社會報告.py")
                    
        except Exception as e:
            st.error(f"❌ 生成失敗：{e}")
            st.exception(e)
else:
    st.info("請先完成 Step 1 後再生成公司段報告")

