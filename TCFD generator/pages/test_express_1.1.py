"""
Express 通道測試頁面：1.1 我們的公司
獨立測試通道，不修改原有 content_pptx_company.py
直接從 +1 步驟生成的 150 字分析文件讀取，硬寫入 prompt
"""
import streamlit as st
import sys
import json
from pathlib import Path
from datetime import datetime

# 導入共享模組
sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import *
from shared.utils import render_sidebar_navigation, render_api_key_input

# 頁面配置
st.set_page_config(page_title="Express 通道測試：1.1 我們的公司", page_icon="🧪", layout="wide")

# 側邊欄
render_sidebar_navigation()
st.sidebar.divider()
API_KEY = render_api_key_input()

# 主頁面
st.title("🧪 Express 通道測試：1.1 我們的公司")
st.markdown("**獨立測試通道，不修改原有 content_pptx_company.py**")

st.divider()

# Express 通道函數
def read_industry_analysis_express() -> str:
    """
    Express 通道：直接從 +1 步驟生成的 150 字分析文件讀取（絕對路徑，不抽象）
    只讀取 150 字分析，不抽取產業別
    """
    log_dir = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")
    if not log_dir.exists():
        st.error(f"❌ Log 目錄不存在: {log_dir}")
        return ""
    
    # 直接讀取最新的 industry_analysis.json 文件（+1 步驟生成的）
    industry_analysis_files = sorted(
        log_dir.glob("session_*_industry_analysis.json"),
        key=lambda f: f.stat().st_mtime,
        reverse=True
    )
    
    if not industry_analysis_files:
        st.warning("⚠️ 找不到 industry_analysis.json 文件")
        return ""
    
    # 讀取最新的文件（絕對路徑）
    log_file = industry_analysis_files[0]
    st.info(f"📁 讀取文件: `{log_file.name}`")
    
    try:
        with open(log_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # 只讀取 150 字分析，不抽取產業別
        industry_analysis = data.get("industry_analysis", "").strip()
        
        if industry_analysis and len(industry_analysis) > 50:
            st.success(f"✅ 讀取 150 字分析成功: {len(industry_analysis)}字")
            return industry_analysis
        else:
            st.warning(f"⚠️ {log_file.name} 中沒有有效的 150 字分析")
            return ""
    except Exception as e:
        st.error(f"❌ 讀取 {log_file.name} 失敗: {e}")
        return ""


def generate_cooperation_info_prompt_express(company_name: str = "本公司") -> str:
    """
    Express 通道：生成 1.1 我們的公司 prompt
    直接硬寫入 150 字分析，無條件判斷，不抽取產業別
    """
    # 直接讀取 150 字分析（Express 通道，絕對路徑，不抽象）
    industry_analysis = read_industry_analysis_express()
    
    if not industry_analysis:
        return ""
    
    # 只有一個 prompt，直接硬寫入 150 字分析（無 if/else，無選擇，不抽取產業別）
    prompt = f"""【⚠️ 最高優先級 - 產業別分析（必須嚴格遵守，不可違反）】
以下產業別分析是本次生成的核心基礎，所有內容必須基於此分析，不得偏離：

{industry_analysis}

【任務】
請根據上述產業別分析，撰寫約 345 字（對應 230 英文單字）描述公司的合作概況，用於 ESG 報告。

【⚠️ 強制要求（必須遵守）】
1. 第一句使用 {{COMPANY_NAME}} 作為公司名稱佔位符
2. 【必須】引用上述產業別分析中的具體數據（如年營收、碳排數據、耗能等級等），不得忽略或抽象化
3. 【必須】內容與上述產業別分析完全一致，不得產生矛盾
4. 使用「我們」和「本公司」，保持第一人稱視角
5. 使用簡潔的中文，不使用項目符號，保持高階主管語調

【公司資訊】
公司名稱：{company_name}

【⚠️ 再次提醒】
上述產業別分析是本次生成的核心基礎，所有內容必須基於此分析，不得偏離。"""
    
    return prompt


# 測試界面
st.subheader("📊 測試 Express 通道")

# 公司名稱輸入
company_name = st.text_input("公司名稱（可選）", "本公司", key="test_company_name")

# 讀取 150 字分析
if st.button("🔍 讀取 150 字分析", type="primary"):
    industry_analysis = read_industry_analysis_express()
    
    if industry_analysis:
        st.success(f"✅ 成功讀取 150 字分析（{len(industry_analysis)}字）")
        
        with st.expander("📝 查看 150 字分析內容", expanded=True):
            st.text_area("150 字分析", industry_analysis, height=200, key="analysis_display")
        
        # 生成 prompt
        prompt = generate_cooperation_info_prompt_express(company_name)
        
        if prompt:
            st.success("✅ Prompt 生成成功")
            
            with st.expander("📋 查看生成的 Prompt", expanded=True):
                st.text_area("Prompt 內容", prompt, height=400, key="prompt_display")
            
            # 下載 prompt
            st.download_button(
                label="📥 下載 Prompt",
                data=prompt,
                file_name=f"express_prompt_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )
            
            # 測試 LLM 調用（可選）
            if API_KEY and st.button("🚀 測試 LLM 調用", type="secondary"):
                try:
                    import anthropic
                    client = anthropic.Anthropic(api_key=API_KEY)
                    
                    with st.spinner("正在調用 LLM..."):
                        response = client.messages.create(
                            model="claude-sonnet-4-20250514",
                            max_tokens=1000,
                            messages=[{"role": "user", "content": prompt}]
                        )
                        
                        result = response.content[0].text if response.content else ""
                        
                        st.success("✅ LLM 調用成功")
                        st.text_area("LLM 回應", result, height=300, key="llm_result")
                        
                except Exception as e:
                    st.error(f"❌ LLM 調用失敗: {e}")
    else:
        st.error("❌ 無法讀取 150 字分析")

st.divider()
st.markdown("### 📌 說明")
st.info("""
**Express 通道特點：**
1. ✅ 直接從 +1 步驟生成的 150 字分析文件讀取（絕對路徑）
2. ✅ 不抽取產業別，只讀取 150 字分析
3. ✅ 不使用環境段 log（不抽象）
4. ✅ 直接硬寫入 prompt，無條件判斷
5. ✅ 獨立測試，不修改原有 content_pptx_company.py
""")

