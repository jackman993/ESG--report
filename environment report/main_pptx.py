"""
ESG 環境篇 PPTX 報告生成器 - 主程式
"""
from environment_pptx import EnvironmentPPTXEngine
from config import ENVIRONMENT_CONFIG
from datetime import datetime
import os


def main(test_mode=False):
    """
    主程式
    test_mode: True = 跳過 API 呼叫，快速測試版面
    """
    
    # 確保 output 資料夾存在
    os.makedirs("output", exist_ok=True)
    
    # 模板路徑（可選）
    template_path = r"C:\Users\User\Downloads\templet_resaved.pptx"
    
    # 檢查模板是否存在
    if os.path.exists(template_path):
        print(f"使用模板：{template_path}")
    else:
        print("未找到模板，使用預設版面")
        template_path = None
    
    # 生成環境篇報告
    engine = EnvironmentPPTXEngine(template_path=template_path, test_mode=test_mode)
    report = engine.generate()
    
    # 儲存檔案
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"output/ESG環境篇_{timestamp}.pptx"
    
    engine.save(filename)
    
    print(f"\n✓ 環境篇 PPTX 報告已儲存：{filename}")
    print(f"✓ 章節：{ENVIRONMENT_CONFIG['chapter_title']}")
    print(f"✓ 總頁數：{len(report.slides)} 頁")
    print("\n完成！")


if __name__ == "__main__":
    import sys
    # 使用 --test 參數啟用測試模式
    test_mode = "--test" in sys.argv
    if test_mode:
        print("=" * 50)
        print("🧪 測試模式：跳過 Claude API 呼叫")
        print("=" * 50)
    main(test_mode=test_mode)

