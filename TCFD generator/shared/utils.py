"""共享工具函數"""
import streamlit as st
import anthropic
from pathlib import Path
from shared.config import ESG_OUTPUT_ROOT

def switch_page(page_path: str):
    """
    切換頁面的輔助函數（替代 st.switch_page，不依賴 pages 系統）
    
    Args:
        page_path: 目標頁面路徑，例如 "pages/0_🏠_首頁.py"
    """
    st.session_state.current_page = page_path
    st.session_state.page_changed = True
    st.rerun()

def clear_all_data():
    """清除所有前次資料，重置系統狀態"""
    from pathlib import Path
    
    # 清除 session_state 中的完成狀態和所有相關資料
    keys_to_clear = [
        # 步驟完成狀態
        "step1_done", "step2_done", "step3_done", "step4_done",
        "emission_done",
        # 資料相關
        "emission_data", "tcfd_summary", "company_profile", "company_name",
        "industry", "industry_selected", "session_id", "timestamp",
        # 輸出相關
        "step1_output_filename", "step2_output_filename", "step3_output_filename",
        "tcfd_output_folder", "emission_output_folder",
        # 其他可能的狀態變數
        "current_step", "report_generated", "output_path",
        "emission_calculated", "tcfd_generated", "company_report_generated",
        "governance_report_generated", "final_report_generated",
        # 確認狀態
        "confirm_reset",
        # API Key 保留（不清除，讓用戶可以繼續使用）
        # "api_key",  # 註釋掉，保留 API Key
    ]
    
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
    
    # 清除 log 文件（可選，保留註釋以便用戶選擇）
    # 如果需要清除 log 文件，取消以下註釋：
    # log_dir = Path(ESG_OUTPUT_ROOT) / "_Backend" / "user_logs"
    # if log_dir.exists():
    #     for log_file in log_dir.glob("*.json"):
    #         try:
    #             log_file.unlink()
    #         except Exception as e:
    #             st.warning(f"無法刪除 {log_file.name}: {e}")
    
    # 重置到首頁
    st.session_state.current_page = "pages/0_🏠_首頁.py"
    
    return True

def render_sidebar_navigation():
    """渲染側邊欄導航（使用按鈕，不依賴 pages 系統）"""
    # 隱藏 Streamlit 自動導航（上半部），保留自定義導航
    st.sidebar.markdown("""
    <style>
        /* 隱藏 Streamlit 自動生成的導航（側邊欄頂部） */
        [data-testid="stSidebarNav"] {
            display: none !important;
        }
    </style>
    """, unsafe_allow_html=True)
    
    st.sidebar.markdown("### 📋 導航")
    
    pages = [
        ("🏠 首頁", "pages/0_🏠_首頁.py"),
        ("🌍 Step 1: 碳排與TCFD氣候治理", "pages/1_🌍_碳排與TCFD氣候治理.py"),
        ("📋 Step 2: 重大議題與公司段報告", "pages/4_📋_重大議題段報告.py"),
        ("🏛️ Step 3: 治理與社會段報告", "pages/5_🏛️_治理與社會報告.py"),
        ("📚 Step 4: 彙整總報告", "pages/6_📚_彙整總報告.py"),
    ]
    
    # 使用按鈕替代 page_link，避免依賴 pages 系統
    for label, page in pages:
        # 檢查是否為當前頁面（用於視覺反饋）
        current_page = st.session_state.get("current_page", "pages/0_🏠_首頁.py")
        is_current = (current_page == page)
        
        # 使用 button 替代 page_link
        if st.sidebar.button(
            label, 
            use_container_width=True, 
            key=f"nav_{page}",
            type="primary" if is_current else "secondary"
        ):
            # 設定目標頁面到 session_state，然後重新載入
            st.session_state.current_page = page
            st.session_state.page_changed = True
            st.rerun()
    
    # 添加重新開始按鈕
    st.sidebar.divider()
    st.sidebar.markdown("### 🔄 重新開始")
    
    # 檢查是否處於確認狀態
    if st.session_state.get("confirm_reset", False):
        st.sidebar.warning("⚠️ 確定要清除所有資料並重新開始嗎？")
        col1, col2 = st.sidebar.columns(2)
        with col1:
            if st.sidebar.button("✅ 確認", use_container_width=True, key="confirm_clear"):
                clear_all_data()
                st.session_state.confirm_reset = False
                st.sidebar.success("✅ 資料已清除！")
                st.rerun()
        with col2:
            if st.sidebar.button("❌ 取消", use_container_width=True, key="cancel_clear"):
                st.session_state.confirm_reset = False
                st.rerun()
    else:
        # 顯示重新開始按鈕
        if st.sidebar.button("🔄 清除資料並重新開始", use_container_width=True, type="secondary", key="reset_button"):
            st.session_state.confirm_reset = True
            st.rerun()

def render_output_folder_links():
    """渲染輸出檔案櫃連結"""
    st.sidebar.markdown("### 📁 輸出檔案櫃")
    
    output_folders = {
        "A_TCFD": ESG_OUTPUT_ROOT / "A_TCFD",
        "B_Emission": ESG_OUTPUT_ROOT / "B_Emission",
        "D_Company": ESG_OUTPUT_ROOT / "D_Company",
        "C_Environment": ESG_OUTPUT_ROOT / "C_Environment",
        "F_Governance_Social": ESG_OUTPUT_ROOT / "F_Governance_Social",
    }
    
    for name, folder_path in output_folders.items():
        if folder_path.exists():
            # 使用 file:// 協議打開文件夾（Windows）
            folder_url = f"file:///{folder_path.as_posix()}"
            st.sidebar.markdown(f"- [{name}]({folder_url})")
        else:
            st.sidebar.markdown(f"- {name} (尚未建立)")

def render_api_key_input():
    """
    渲染 API Key 輸入
    
    優先順序：
    1. Streamlit Secrets（生產環境，部署後自動讀取）
    2. session_state（開發環境，跨頁面共享）
    3. 用戶輸入（開發模式，fallback）
    """
    st.sidebar.markdown("### ⚙️ 設定")
    
    # 1. 優先從 Streamlit Secrets 讀取（生產環境/朋友試用）
    try:
        if hasattr(st, "secrets") and st.secrets and "api_keys" in st.secrets:
            api_key = st.secrets["api_keys"]["anthropic_key"]
            if api_key and api_key.strip() and api_key != "your-anthropic-api-key-here":
                # 保存到 session_state 以便跨頁面使用
                st.session_state.api_key = api_key.strip()
                st.sidebar.success("✅ API Key 已自動配置")
                # 不顯示輸入框，直接返回（朋友試用時無需輸入）
                return api_key.strip()
    except Exception:
        # secrets 不存在或讀取失敗，繼續下一步
        pass
    
    # 2. 從 session_state 讀取（開發環境，跨頁面共享）
    if "api_key" in st.session_state and st.session_state.api_key:
        api_key = st.session_state.api_key
        st.sidebar.success("✅ API Key 已設置")
        # 顯示清除按鈕（僅在開發模式顯示）
        if st.sidebar.button("🗑️ 清除 API Key", use_container_width=True, key="clear_api_key"):
            st.session_state.api_key = ""
            st.rerun()
        return api_key
    
    # 3. 開發模式：允許輸入（fallback，僅在沒有配置時顯示）
    st.sidebar.info("💡 請輸入 API Key 或配置 secrets.toml")
    
    # 初始化 session_state（如果還沒有）
    if "api_key" not in st.session_state:
        st.session_state.api_key = ""
    
    # 使用臨時的 key 來輸入，避免與 session_state 衝突
    input_key = st.sidebar.text_input(
        "🔑 Claude API Key", 
        type="password", 
        key="api_key_input_temp",  # 使用臨時 key，避免與 session_state 衝突
        value="",  # 不預填，避免顯示已保存的值
        placeholder="貼上 API Key 後點擊「確認保存」"
    )
    
    # 確認按鈕
    if st.sidebar.button("✅ 確認保存", use_container_width=True, type="primary", key="confirm_api_key"):
        if input_key and len(input_key.strip()) > 0:
            st.session_state.api_key = input_key.strip()
            st.sidebar.success("✅ API Key 已保存！")
            st.rerun()
        else:
            st.sidebar.error("❌ 請先輸入 API Key")
    
    # 返回已保存的 key（優先）或輸入的 key（臨時使用）
    api_key = st.session_state.get("api_key") or input_key
    
    if not api_key or not api_key.strip():
        st.sidebar.warning("⚠️ 請輸入 API Key 並點擊「確認保存」")
        return ""  # 返回空字串而不是 None
    
    # 清理 API key（移除前後空白）
    api_key = api_key.strip()
    
    return api_key

def generate_report_summary(step: str, context_data: dict, api_key: str, test_mode: bool = False) -> str:
    """
    生成報告摘要（200字）
    
    Args:
        step: 步驟名稱（"Step 1", "Step 2", "Step 3"）
        context_data: 上下文數據（產業、公司資料、TCFD摘要等）
        api_key: Claude API key
        test_mode: 測試模式（跳過LLM調用）
    
    Returns:
        200字摘要
    """
    if test_mode:
        return "【測試模式】報告摘要：本報告涵蓋環境治理、碳排放管理、TCFD氣候風險評估等關鍵議題，展現公司在永續發展方面的具體作為與成果。"
    
    # 驗證 API key
    if not api_key or not api_key.strip():
        return "❌ API Key 未設置，無法生成摘要"
    
    try:
        client = anthropic.Anthropic(api_key=api_key.strip())
        
        # 根據步驟構建不同的prompt
        if step == "Step 1":
            # 環境段摘要
            industry = context_data.get("industry", "企業")
            company_profile = context_data.get("company_profile", {})
            emission_data = context_data.get("emission_data", {})
            tcfd_summary = context_data.get("tcfd_summary", {})
            session_id = context_data.get("session_id", "")
            
            # 讀取 150 字產業別分析（硬插入）
            industry_analysis = ""
            if session_id:
                try:
                    import json
                    from pathlib import Path
                    log_dir = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")
                    log_file = log_dir / f"session_{session_id}_industry_analysis.json"
                    
                    if log_file.exists():
                        with open(log_file, "r", encoding="utf-8") as f:
                            data = json.load(f)
                        industry_analysis = data.get("industry_analysis", "").strip()
                        if industry_analysis:
                            print(f"[摘要生成] 讀取到 150 字分析: {len(industry_analysis)}字")
                except Exception as e:
                    print(f"[WARN] 讀取 150 字分析失敗: {e}")
            
            # 150字硬切入 prompt 最前面
            if industry_analysis:
                prompt = f"""【硬性要求 - 產業別分析（必須嚴格遵守）】
{industry_analysis}

【任務】
請根據上述產業別分析，生成200字ESG環境段報告摘要。

【要求】
1. 必須引用上述產業別分析中的具體數據（如年營收、碳排數據、耗能等級等）
2. 內容必須與上述產業別分析一致
3. 重點說明：
   - 公司的環境治理架構
   - 碳排放管理策略
   - TCFD氣候風險因應措施
   - 永續發展目標與成果

【補充資訊】
**產業類別：** {industry}
**公司規模：** {company_profile.get('size', '未知')}
**年度營收：** {company_profile.get('annual_revenue_wan', '未知')}萬元
**碳排放數據：** {emission_data.get('total', emission_data.get('total_emission', '未提供'))}
**TCFD轉型風險：** {tcfd_summary.get('transformation_policy', '未提供')[:100] if tcfd_summary.get('transformation_policy') else '未提供'}
**TCFD市場風險：** {tcfd_summary.get('market_trend', '未提供')[:100] if tcfd_summary.get('market_trend') else '未提供'}

摘要要求：精簡、專業、突出重點，約200字。**重要：請使用純文本格式，不要使用 Markdown 標題（如 #、##）或任何格式符號，直接輸出摘要文字即可。**"""
            else:
                # 如果沒有150字分析，使用原來的 prompt
                prompt = f"""請為以下ESG環境段報告生成200字摘要：

**產業類別：** {industry}
**公司規模：** {company_profile.get('size', '未知')}
**年度營收：** {company_profile.get('annual_revenue_wan', '未知')}萬元

**碳排放數據：**
{emission_data.get('total', emission_data.get('total_emission', '未提供'))}

**TCFD氣候風險摘要：**
- 轉型風險：{tcfd_summary.get('transformation_policy', '未提供')[:100] if tcfd_summary.get('transformation_policy') else '未提供'}
- 市場風險：{tcfd_summary.get('market_trend', '未提供')[:100] if tcfd_summary.get('market_trend') else '未提供'}

請生成200字摘要，重點說明：
1. 公司的環境治理架構
2. 碳排放管理策略
3. TCFD氣候風險因應措施
4. 永續發展目標與成果

摘要要求：精簡、專業、突出重點，約200字。**重要：請使用純文本格式，不要使用 Markdown 標題（如 #、##）或任何格式符號，直接輸出摘要文字即可。**"""
        
        elif step == "Step 2":
            # 公司段摘要
            company_name = context_data.get("company_name", "本公司")
            
            prompt = f"""請為以下ESG重大議題與公司段報告生成200字摘要：

**公司名稱：** {company_name}

本報告涵蓋：
- 重大議題分析
- 公司永續策略
- 利害關係人溝通
- 永續發展目標與績效

請生成200字摘要，重點說明：
1. 公司識別的重大永續議題
2. 永續策略與管理方針
3. 利害關係人溝通機制
4. 具體成果與未來規劃

摘要要求：精簡、專業、突出重點，約200字。**重要：請使用純文本格式，不要使用 Markdown 標題（如 #、##）或任何格式符號，直接輸出摘要文字即可。**"""
        
        elif step == "Step 3":
            # 治理與社會段摘要
            prompt = f"""請為以下ESG治理與社會段報告生成200字摘要：

本報告涵蓋：
- 公司治理架構與運作
- 董事會職能與監督機制
- 風險管理與內控制度
- 社會責任與員工權益
- 社區參與與社會貢獻

請生成200字摘要，重點說明：
1. 公司治理架構與運作機制
2. 風險管理與內控制度
3. 員工權益與職場環境
4. 社會責任與社區參與

摘要要求：精簡、專業、突出重點，約200字。**重要：請使用純文本格式，不要使用 Markdown 標題（如 #、##）或任何格式符號，直接輸出摘要文字即可。**"""
        
        else:
            return "摘要生成中..."
        
        # 調用Claude API
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=300,
            messages=[{
                "role": "user",
                "content": prompt
            }]
        )
        
        summary = response.content[0].text.strip()
        
        # 清理 Markdown 格式（移除標題符號，避免顯示為大標題）
        import re
        # 移除 Markdown 標題符號（#、##、### 等）
        summary = re.sub(r'^#+\s*', '', summary, flags=re.MULTILINE)
        # 移除其他常見的 Markdown 格式符號
        summary = re.sub(r'\*\*([^*]+)\*\*', r'\1', summary)  # 移除粗體
        summary = re.sub(r'\*([^*]+)\*', r'\1', summary)  # 移除斜體
        
        # 確保摘要約200字（如果太長則截斷）
        if len(summary) > 250:
            summary = summary[:250] + "..."
        
        return summary
    
    except anthropic.AuthenticationError as auth_err:
        # API 認證錯誤
        error_msg = str(auth_err)
        if "redacted" in error_msg.lower() or "api key" in error_msg.lower():
            return "❌ API Key 認證失敗：請檢查 API Key 是否正確或已過期。前往 https://console.anthropic.com/ 確認 API Key 狀態"
        return f"❌ API 認證失敗：{error_msg[:100]}"
    except anthropic.APIError as api_err:
        # API 調用錯誤（配額、服務不可用等）
        return f"❌ API 調用失敗：可能是配額用盡或服務暫時不可用，請稍後再試"
    except Exception as e:
        # 其他錯誤
        error_msg = str(e)
        if len(error_msg) > 100:
            error_msg = error_msg[:100] + "..."
        return f"❌ 摘要生成失敗：{error_msg}"

