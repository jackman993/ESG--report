"""
ESG 報告生成系統 - Streamlit Cloud 入口點
使用虛擬上層 pages/ 橋接器，轉發到 TCFD generator 的實際頁面
"""
# Trigger redeploy - 觸發重新部署
import sys
import os
from pathlib import Path
import streamlit as st

# 添加 TCFD generator 路徑到 Python 路徑（讓橋接器可以找到實際頁面）
tcfd_path = Path(__file__).parent / "TCFD generator"
if tcfd_path.exists():
    sys.path.insert(0, str(tcfd_path))

# 初始化 session_state（如果還沒有）
if "current_page" not in st.session_state:
    st.session_state.current_page = "pages/0_🏠_首頁.py"

# 根據 session_state 動態載入對應的頁面
import importlib.util
target_page = st.session_state.current_page

# 將相對路徑轉換為實際檔案路徑
# 例如 "pages/0_🏠_首頁.py" -> 虛擬上層的 pages/0_🏠_首頁.py
if target_page.startswith("pages/"):
    page_filename = target_page.replace("pages/", "")
    pages_path = Path(__file__).parent / "pages" / page_filename
else:
    # 如果路徑不正確，回退到首頁
    pages_path = Path(__file__).parent / "pages" / "0_🏠_首頁.py"
    st.session_state.current_page = "pages/0_🏠_首頁.py"

if pages_path.exists():
    spec = importlib.util.spec_from_file_location("page_module", str(pages_path))
    if spec and spec.loader:
        page_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(page_module)
    else:
        st.error(f"無法載入頁面模組: {pages_path}")
        # 回退到首頁
        st.session_state.current_page = "pages/0_🏠_首頁.py"
        st.rerun()
else:
    st.error(f"找不到頁面文件: {pages_path}")
    st.info(f"當前頁面: {target_page}")
    st.info(f"TCFD generator 路徑: {tcfd_path}")
    # 回退到首頁
    st.session_state.current_page = "pages/0_🏠_首頁.py"
    if tcfd_path.exists():
        tcfd_pages_dir = tcfd_path / "pages"
        if tcfd_pages_dir.exists():
            st.info(f"TCFD generator pages 目錄存在，包含文件:")
            for f in sorted(tcfd_pages_dir.glob("*.py")):
                st.text(f"  - {f.name}")
