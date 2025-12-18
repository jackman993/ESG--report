"""
公司段引擎包裝器（中文版）- 方案一：簡單包裝器
不修改現有引擎，通過環境變數和臨時修改 config 來傳遞參數
"""
import os
import sys
from pathlib import Path
from typing import Tuple, Optional

def generate_company_section_zh(
    api_key: str,
    company_name: str = None,
    output_dir: Path = None
) -> Tuple[Optional[str], Optional[str]]:
    """
    生成中文版公司段 PPTX（方案一：簡單包裝器）
    
    Args:
        api_key: Claude API key
        company_name: 公司名稱（可選）
        output_dir: 輸出目錄（可選，預設使用 config 中的路徑）
    
    Returns:
        (輸出文件路徑, 錯誤訊息) - 成功時返回 (路徑, None)，失敗時返回 (None, 錯誤訊息)
    """
    # 保存原始狀態
    original_path = sys.path.copy()
    
    try:
        # 1. 設置引擎路徑（從 ESG go 目錄）
        BASE_DIR = Path(__file__).parent.parent  # ESG go/
        company_path = BASE_DIR / "company1.1-3.6"
        if not company_path.exists():
            return None, f"引擎路徑不存在: {company_path}"
        
        sys.path.insert(0, str(company_path))
        
        # 2. 導入 config
        import config_pptx_company
        
        # 3. 臨時修改 config 中的 API key
        original_key = config_pptx_company.ANTHROPIC_API_KEY
        config_pptx_company.ANTHROPIC_API_KEY = api_key
        
        # 4. 如果指定了輸出目錄，臨時修改 config
        original_output = None
        if output_dir:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            original_output = config_pptx_company.OUTPUT_PATH
            config_pptx_company.OUTPUT_PATH = str(output_dir)
        
        # 5. 導入引擎（會從 config 讀取 API key）
        from content_pptx_company import PPTContentEngine
        from full_pptx_company import PPTFullEngine
        
        # 6. 初始化（會自動讀取環境段 log）
        content_engine = PPTContentEngine()  # 從 config 讀取 API key
        ppt_engine = PPTFullEngine(content_engine, company_name=company_name)
        
        # 7. 生成 PPTX
        output_path = ppt_engine.generate(company_name=company_name)
        
        return (output_path, None)
    
    except Exception as e:
        import traceback
        error_msg = f"生成公司段失敗: {str(e)}\n{traceback.format_exc()}"
        return (None, error_msg)
    
    finally:
        # 8. 恢復 config
        if 'config_pptx_company' in sys.modules:
            import config_pptx_company
            if 'original_key' in locals():
                config_pptx_company.ANTHROPIC_API_KEY = original_key
            if 'original_output' in locals() and original_output:
                config_pptx_company.OUTPUT_PATH = original_output
        
        # 恢復 sys.path
        sys.path[:] = original_path

