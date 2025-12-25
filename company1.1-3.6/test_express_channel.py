"""
Express 通道測試：1.1 我們的公司
直接從 +1 步驟生成的 150 字分析文件讀取，硬寫入 prompt
不抽取產業別，不使用環境段 log
"""
import json
from pathlib import Path
from typing import Dict, Optional


def read_industry_analysis_express() -> str:
    """
    Express 通道：直接從 +1 步驟生成的 150 字分析文件讀取（絕對路徑，不抽象）
    只讀取 150 字分析，不抽取產業別
    """
    log_dir = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")
    if not log_dir.exists():
        print(f"[Express] Log 目錄不存在: {log_dir}")
        return ""
    
    # 直接讀取最新的 industry_analysis.json 文件（+1 步驟生成的）
    industry_analysis_files = sorted(
        log_dir.glob("session_*_industry_analysis.json"),
        key=lambda f: f.stat().st_mtime,
        reverse=True
    )
    
    if not industry_analysis_files:
        print(f"[Express] 找不到 industry_analysis.json 文件")
        return ""
    
    # 讀取最新的文件（絕對路徑）
    log_file = industry_analysis_files[0]
    print(f"[Express] 讀取文件: {log_file}")
    
    try:
        with open(log_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # 只讀取 150 字分析，不抽取產業別
        industry_analysis = data.get("industry_analysis", "").strip()
        
        if industry_analysis and len(industry_analysis) > 50:
            print(f"[Express] 讀取 150 字分析成功: {len(industry_analysis)}字")
            print(f"[Express] 前100字: {industry_analysis[:100]}...")
            return industry_analysis
        else:
            print(f"[Express] {log_file.name} 中沒有有效的 150 字分析")
            return ""
    except Exception as e:
        print(f"[Express] 讀取 {log_file.name} 失敗: {e}")
        return ""


def generate_cooperation_info_prompt_express(company_name: str = "本公司") -> str:
    """
    Express 通道：生成 1.1 我們的公司 prompt
    直接硬寫入 150 字分析，無條件判斷，不抽取產業別
    """
    # 直接讀取 150 字分析（Express 通道，絕對路徑，不抽象）
    industry_analysis = read_industry_analysis_express()
    
    # 調試：檢查實際值
    print(f"\n{'='*60}")
    print(f"[Express generate_cooperation_info] 150字分析長度={len(industry_analysis) if industry_analysis else 0}")
    print(f"[Express generate_cooperation_info] 150字分析前100字={industry_analysis[:100] if industry_analysis else 'None'}")
    print(f"{'='*60}\n")
    
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
    
    print(f"[Express OK] 150字硬寫入 prompt（{len(industry_analysis) if industry_analysis else 0}字），無條件判斷，只有一個 prompt，不抽取產業別")
    
    return prompt


if __name__ == "__main__":
    print("=" * 60)
    print("Express 通道測試：1.1 我們的公司")
    print("=" * 60)
    
    # 測試讀取 150 字分析
    industry_analysis = read_industry_analysis_express()
    
    if industry_analysis:
        print(f"\n✅ 成功讀取 150 字分析（{len(industry_analysis)}字）")
        
        # 測試生成 prompt
        prompt = generate_cooperation_info_prompt_express()
        
        print(f"\n{'='*60}")
        print("生成的 Prompt 內容：")
        print(f"{'='*60}")
        print(prompt)
        print(f"{'='*60}\n")
    else:
        print("\n❌ 無法讀取 150 字分析，請確認 +1 步驟已生成 industry_analysis.json 文件")

