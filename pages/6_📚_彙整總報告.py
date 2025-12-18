"""
虛擬上層橋接器 - 轉發到 TCFD generator 的實際頁面
"""
import sys
import os
from pathlib import Path
import importlib.util

# 取得 TCFD generator 的實際頁面路徑
base_path = Path(__file__).parent.parent
tcfd_pages_path = base_path / "TCFD generator" / "pages" / "6_📚_彙整總報告.py"

# 將 TCFD generator 添加到 Python 路徑
tcfd_path = base_path / "TCFD generator"
sys.path.insert(0, str(tcfd_path))

# 切換工作目錄到 TCFD generator
original_cwd = os.getcwd()
os.chdir(str(tcfd_path))

try:
    # 載入並執行實際的頁面模組
    if tcfd_pages_path.exists():
        spec = importlib.util.spec_from_file_location("real_page", str(tcfd_pages_path))
        if spec and spec.loader:
            real_page = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(real_page)
        else:
            import streamlit as st
            st.error(f"無法載入頁面模組: {tcfd_pages_path}")
    else:
        import streamlit as st
        st.error(f"找不到頁面文件: {tcfd_pages_path}")
finally:
    os.chdir(original_cwd)

