"""
TCFD generator/app.py - 重新導向到根目錄的 app.py
此文件用於解決 Streamlit Cloud 路徑配置問題

如果 Streamlit Cloud 配置為在 TCFD generator/ 目錄尋找 app.py，
此文件會重新導向到根目錄的實際 app.py
"""
import sys
import os
from pathlib import Path
import importlib.util

# 獲取根目錄的 app.py 路徑
current_dir = Path(__file__).parent
root_dir = current_dir.parent
root_app = root_dir / "app.py"

# 如果根目錄的 app.py 存在，則載入並執行它
if root_app.exists():
    # 將根目錄添加到 Python 路徑
    if str(root_dir) not in sys.path:
        sys.path.insert(0, str(root_dir))
    
    # 使用 importlib 載入根目錄的 app.py 模組
    spec = importlib.util.spec_from_file_location("root_app", str(root_app))
    if spec and spec.loader:
        # 設置 __file__ 為根目錄的 app.py，這樣相對路徑才能正確解析
        module = importlib.util.module_from_spec(spec)
        module.__file__ = str(root_app)
        spec.loader.exec_module(module)
    else:
        import streamlit as st
        st.error(f"無法載入根目錄的 app.py: {root_app}")
else:
    import streamlit as st
    st.error(f"找不到根目錄的 app.py: {root_app}")
    st.info("請確認 Streamlit Cloud 的主文件路徑設置為根目錄的 app.py")

