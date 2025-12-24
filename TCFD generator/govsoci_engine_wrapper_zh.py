"""
治理社會段引擎包裝器（中文版）- 方案一：簡單包裝器
不修改現有引擎，通過環境變數和臨時修改 config 來傳遞參數
"""
import os
import sys
import logging
from pathlib import Path
from typing import Tuple, Optional

# 設置日誌
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    ))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

def generate_govsoci_section_zh(
    api_key: str,
    output_dir: Path = None
) -> Tuple[Optional[str], Optional[str]]:
    """
    生成中文版治理社會段 PPTX（方案一：簡單包裝器）
    
    Args:
        api_key: Claude API key
        output_dir: 輸出目錄（可選，預設使用 config 中的路徑）
    
    Returns:
        (輸出文件路徑, 錯誤訊息) - 成功時返回 (路徑, None)，失敗時返回 (None, 錯誤訊息)
    """
    # ========== 驗證輸入 ==========
    if not api_key or not api_key.strip():
        logger.error("API Key 驗證失敗：API Key 不能為空")
        return None, "API Key 不能為空"
    
    if not api_key.startswith("sk-ant-"):
        logger.warning(f"API Key 格式可能不正確（應以 sk-ant- 開頭）")
    
    # ========== 保存原始狀態 ==========
    original_path = sys.path.copy()
    original_key = None
    original_output = None
    config_module = None
    
    logger.info("開始生成治理與社會段 PPTX")
    logger.debug(f"原始 sys.path 長度: {len(original_path)}")
    
    try:
        # 1. 設置引擎路徑（從 ESG go 目錄）
        BASE_DIR = Path(__file__).parent.parent  # ESG go/
        govsoci_path = BASE_DIR / "GovSoci5.1-6.9"
        
        logger.debug(f"檢查引擎路徑: {govsoci_path}")
        if not govsoci_path.exists():
            logger.error(f"引擎路徑不存在: {govsoci_path}")
            return None, f"引擎路徑不存在: {govsoci_path}"
        
        logger.info(f"載入治理與社會段引擎: {govsoci_path}")
        sys.path.insert(0, str(govsoci_path))
        logger.debug(f"更新後 sys.path 長度: {len(sys.path)}")
        
        # 2. 導入 config
        import config_pptx
        config_module = config_pptx
        logger.debug("成功導入 config_pptx")
        
        # 3. 臨時修改 config 中的 API key
        original_key = getattr(config_pptx, 'ANTHROPIC_API_KEY', None)
        config_pptx.ANTHROPIC_API_KEY = api_key.strip()
        logger.debug("已設置 API Key 到 config")
        
        # 4. 如果指定了輸出目錄，臨時修改 config
        if output_dir:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            original_output = getattr(config_pptx, 'OUTPUT_PATH', None)
            config_pptx.OUTPUT_PATH = str(output_dir)
            logger.info(f"設置輸出目錄: {output_dir}")
        
        # 5. 導入引擎（會從 config 讀取 API key）
        logger.debug("導入引擎模組...")
        from content_pptx import PPTContentEngine
        from full_pptx import PPTFullEngine
        logger.debug("引擎模組導入成功")
        
        # 6. 初始化（會自動讀取環境段 log）
        logger.info("初始化內容引擎...")
        content_engine = PPTContentEngine()  # 從 config 讀取 API key
        logger.info("初始化 PPT 引擎...")
        ppt_engine = PPTFullEngine(content_engine)
        
        # 7. 生成 PPTX
        logger.info("開始生成 PPTX...")
        output_path = ppt_engine.generate()
        logger.info(f"PPTX 生成成功: {output_path}")
        
        return (output_path, None)
    
    except Exception as e:
        import traceback
        error_msg = f"生成治理社會段失敗: {str(e)}"
        logger.error(f"{error_msg}\n{traceback.format_exc()}")
        return (None, error_msg)
    
    finally:
        # 8. 恢復 config（確保執行）
        logger.debug("開始恢復環境狀態...")
        
        # 恢復 API Key
        try:
            if config_module and original_key is not None:
                config_module.ANTHROPIC_API_KEY = original_key
                logger.debug("已恢復 API Key")
            elif config_module:
                logger.warning("無法恢復 API Key（原始值未保存）")
        except Exception as e:
            logger.error(f"恢復 API Key 時發生錯誤: {e}")
        
        # 恢復輸出路徑
        try:
            if config_module and original_output is not None:
                config_module.OUTPUT_PATH = original_output
                logger.debug("已恢復輸出路徑")
        except Exception as e:
            logger.error(f"恢復輸出路徑時發生錯誤: {e}")
        
        # 恢復 sys.path（確保執行）
        try:
            sys.path[:] = original_path
            logger.debug(f"已恢復 sys.path（長度: {len(sys.path)}）")
        except Exception as e:
            logger.error(f"恢復 sys.path 時發生錯誤: {e}")
        
        logger.info("環境狀態恢復完成")

