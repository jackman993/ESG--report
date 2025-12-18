"""
ESG 環境篇報告生成器 - 主程式
"""
from environment_full import EnvironmentFullEngine
from config import ENVIRONMENT_CONFIG
from datetime import datetime


def main():
    """主程式"""

    # 生成環境篇報告
    engine = EnvironmentFullEngine()
    report = engine.generate()

    # 儲存檔案
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"ESG環境篇_{ENVIRONMENT_CONFIG.get('company_name', '企業')}_{timestamp}.docx"

    report.save(filename)

    print(f"\n✓ 環境篇報告已儲存：{filename}")
    print(f"✓ 章節：{ENVIRONMENT_CONFIG['chapter_title']}")
    print(f"✓ 總頁數：{ENVIRONMENT_CONFIG['total_pages']} 頁")
    print("\n完成！")


if __name__ == "__main__":
    main()