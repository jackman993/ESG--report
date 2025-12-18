"""主頁面 - ESG 報告生成系統"""
import streamlit as st
import sys
from pathlib import Path

# 導入共享模組
sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import *
from shared.utils import render_output_folder_links, render_api_key_input, render_sidebar_navigation, switch_page

# 頁面配置
st.set_page_config(page_title="ESG 報告生成系統", page_icon="🏠", layout="wide")

# 側邊欄（自定義導航）
render_sidebar_navigation()
st.sidebar.divider()
API_KEY = render_api_key_input()
render_output_folder_links()

# 主頁面
st.title("🏠 ESG 報告生成系統")
st.markdown("---")

# 整體流程概覽
st.subheader("📋 生成流程")
st.info("""
**完整流程：**
1. **Step 1**: 碳排與TCFD氣候治理
   - 子步驟1: 碳排放計算
   - 子步驟2: TCFD 表格生成
   - 子步驟3: 生成環境治理報告（第四章）- 環境段

2. **Step 2**: 重大議題與公司段報告

3. **Step 3**: 治理與社會段報告

4. **Step 4**: 彙整總報告
""")

st.divider()

# 各步驟完成狀態
st.subheader("✅ 完成狀態")

col1, col2, col3, col4 = st.columns(4)

with col1:
    step1_done = st.session_state.get("step1_done", False)
    if step1_done:
        st.success("✅ Step 1")
    else:
        st.info("⬜ Step 1")

with col2:
    step2_done = st.session_state.get("step2_done", False)
    if step2_done:
        st.success("✅ Step 2")
    else:
        st.info("⬜ Step 2")

with col3:
    step3_done = st.session_state.get("step3_done", False)
    if step3_done:
        st.success("✅ Step 3")
    else:
        st.info("⬜ Step 3")

with col4:
    step4_done = st.session_state.get("step4_done", False)
    if step4_done:
        st.success("✅ Step 4")
    else:
        st.info("⬜ Step 4")

st.divider()

# 快速導航
st.subheader("🚀 快速開始")
st.markdown("""
請從左側導航選擇步驟，或點擊下方按鈕快速開始：
""")

col1, col2 = st.columns(2)
with col1:
    if st.button("🌍 開始 Step 1", use_container_width=True, type="primary"):
        switch_page("pages/1_🌍_碳排與TCFD氣候治理.py")
with col2:
    if st.button("🏭 TCFD與環境段報告", use_container_width=True):
        switch_page("pages/2_🏭_TCFD報告生成器.py")

