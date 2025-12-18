"""Step 3: 治理與社會段報告"""
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
st.set_page_config(page_title="Step 3: 治理與社會段報告", page_icon="🏛️", layout="wide")

# 側邊欄（自定義導航）
render_sidebar_navigation()
st.sidebar.divider()
API_KEY = render_api_key_input()
render_output_folder_links()

# 主頁面
st.title("🏛️ Step 3: 治理與社會段報告")

# 前置條件檢查
st.subheader("📋 前置條件檢查")
step1_done = st.session_state.get("step1_done", False)
step2_done = st.session_state.get("step2_done", False)

col1, col2 = st.columns(2)
with col1:
    st.success("✅ Step 1") if step1_done else st.warning("⬜ Step 1")
with col2:
    st.success("✅ Step 2") if step2_done else st.info("ℹ️ Step 2（可選）")

st.divider()

# 治理與社會段生成
if step1_done:
    st.subheader("👥 治理與社會段生成")
    st.info("生成治理段（5.x）和社會段（6.x）")
    
    # 檢查是否已經生成過（持久化顯示）
    if "step3_output_path" in st.session_state:
        output_path = Path(st.session_state.step3_output_path)
        if output_path.exists():
            st.markdown("### ✅ 已生成的報告")
            
            # 顯示摘要
            if "step3_summary" in st.session_state:
                st.markdown("### 📝 報告摘要")
                st.info(st.session_state.step3_summary)
            
            st.info(f"📁 **完整路徑：** `{output_path}`")
            
            # 下載按鈕
            with open(output_path, "rb") as f:
                file_data = f.read()
            
            st.download_button(
                label="📥 下載治理與社會段 PPTX",
                data=file_data,
                file_name=st.session_state.get("step3_output_filename", output_path.name),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True
            )
            
            st.success(f"✅ 治理與社會段生成完成！")
            
            # 下一步按鈕
            st.divider()
            st.markdown("### 🎯 下一步")
            if st.button("➡️ 下一步：彙整總報告", use_container_width=True, type="primary", key="next_to_step4_persist"):
                switch_page("pages/6_📚_彙整總報告.py")
    
    if st.button("🚀 生成治理與社會段 PPTX", type="primary", use_container_width=True, key="btn_govsoci"):
        if not API_KEY:
            st.error("請先在左側輸入 API Key")
            st.stop()
        
        with st.spinner("生成中...請稍候（約 2-3 分鐘）"):
            try:
                # 導入包裝器
                from govsoci_engine_wrapper_zh import generate_govsoci_section_zh
                
                st.info("📄 正在調用治理與社會段引擎...")
                
                # 調用包裝器
                output_path, error = generate_govsoci_section_zh(
                    api_key=API_KEY,
                    output_dir=OUTPUT_F_GOVSOCI
                )
                
                if error:
                    st.error(f"❌ {error}")
                    st.exception(Exception(error))
                else:
                    # 生成摘要
                    context_data = {}
                    summary = generate_report_summary("Step 3", context_data, API_KEY, False)
                    
                    # 保存到 session_state（持久化）
                    st.session_state.step3_output_path = str(output_path)
                    st.session_state.step3_summary = summary
                    st.session_state.step3_output_filename = Path(output_path).name
                    
                    st.success(f"✅ 治理與社會段生成完成！")
                    st.info(f"📁 **完整路徑：** `{output_path}`")
                    
                    # 顯示摘要
                    st.markdown("### 📝 報告摘要")
                    st.info(summary)
                    
                    st.session_state.step3_done = True
                    
                    # 下載按鈕
                    if Path(output_path).exists():
                        with open(output_path, "rb") as f:
                            file_data = f.read()
                        
                        st.download_button(
                            label="📥 下載治理與社會段 PPTX",
                            data=file_data,
                            file_name=Path(output_path).name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            use_container_width=True
                        )
                        st.balloons()
                        
                        # 保存 session log
                        session_log = {
                            "step": "Step 3",
                            "output_path": str(output_path),
                            "summary": summary
                        }
                        save_session_log(session_log)
                        
                        # 下一步按鈕
                        st.divider()
                        st.markdown("### 🎯 下一步")
                        if st.button("➡️ 下一步：彙整總報告", use_container_width=True, type="primary", key="next_to_step4"):
                            switch_page("pages/6_📚_彙整總報告.py")
                    
            except Exception as e:
                st.error(f"❌ 生成失敗：{e}")
                st.exception(e)
else:
    st.info("請先完成 Step 1 後再生成治理與社會段報告")

