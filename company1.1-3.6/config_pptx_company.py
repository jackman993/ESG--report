"""
PPT configuration for the company section (independent engine build).
"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# API settings (reuse existing keys/models)
ANTHROPIC_API_KEY = "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"
CLAUDE_MODEL = "claude-3-5-sonnet-20241022"
CLAUDE_MODEL_FALLBACKS = [
    "claude-3-sonnet-20240229",
    "claude-3-haiku-20240307",
    "claude-3-opus-20240229",
]
LLM_WORD_COUNT = 280

# Paths
DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")

# 改用使用者在桌面手繪並由 PowerPoint 另存的新模板，避免舊 PptxGenJS 模板的結構問題
SEED_TEMPLATE_PATH = str(BASE_DIR.parent / "assets2" / "handdrawppt.pptx")

OUTPUT_PATH = str(BASE_DIR / "output")
# Assets 已移到上一層目錄
ASSETS2_PATH = str(BASE_DIR.parent / "assets2")
# assets1 在上一層目錄（ESG go\assets1）
ASSETS1_PATH = str(BASE_DIR.parent / "assets1")
ASSETS3_COMPANY_PATH = ASSETS1_PATH  # 保持變數名兼容性
SDG_ICON_FILES = [
    "3-4SDG3.jpg",
    "3-4SDG4.jpg",
    "3-4SDG5.jpg",
    "3-4SDG6.jpg",
    "3-4SDG7.jpg",
    "3-4SDG8.jpg",
    "3-4SDG9.jpg",
    "3-4SDG12.jpg",
    "3-4SDG13.jpg",
]
SDG_ICON_PATHS = [os.path.join(ASSETS3_COMPANY_PATH, name) for name in SDG_ICON_FILES]

# Core PPT settings for experimental build
# 字型策略：
# - 中文：Microsoft JhengHei（微軟正黑體）
# - 英文 / 數字：Calibri（由 PowerPoint 自己 fallback）
# 在程式裡統一設定 font_family 為 "Microsoft JhengHei"，讓所有段落預設用這一套，
# Calibri 會自動作為英文 fallback，避免出現奇怪或自定義字體導致檔案結構不穩定。
PPT_CONFIG = {
    "seed_template": SEED_TEMPLATE_PATH,
    "output_filename": "ESG_PPT_company.pptx",
    "total_slides": 17,
    "font_family": "Microsoft JhengHei",
    "font_size": 11,
    "watermark_text": "Sustainability Report",
}

# Common text snippets
A_TEXT_FALLBACK = (
    "This page uses Layout A, which positions narrative copy in the left column and places a hero image on the right. "
    "The text area remains constrained within its textbox, ensuring the photograph or chart on the right never overlaps the narrative."
)

B_TEXT_FALLBACK = (
    "This page demonstrates Layout B. Supporting imagery is stacked inside the left media frame while the descriptive copy sits on the right. "
    "Each image stays within its bounds so it never collides with the text column."
)

SLIDE_CONFIGS = {
    1: {
        "title": "",
        "layout": "A",
        "paragraphs": 0,
        "image_path": os.path.join(ASSETS3_COMPANY_PATH, "0-0cover_01.png"),
        "image_width_cm": 26.0,
        "image_height_cm": 16.0,
        "image_left_cm": 1.5,
        "image_top_cm": 1.0,
        # Company name text box on cover
        "company_name_box": {
            "left_cm": 2.5,
            "top_cm": 12.0,
            "width_cm": 25.0,
            "height_cm": 2.0,
        },
        "company_name_font_size": 32,
        "company_name_font_color": (255, 255, 255),  # White
        "company_name_bold": True,
    },
    2: {
        "title": "執行長訊息",
        "layout": "A",
        "content_method": "generate_ceo_message",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "image_path": os.path.join(ASSETS3_COMPANY_PATH, "1-1ceom02..png"),
        "image_width_cm": 10.0,
        "image_align": "right",
        "image_top_cm": 3.0,
        "image_left_cm": 18.0,
    },
    3: {
        "title": "1.1 我們的公司",
        "layout": "A",
        "content_method": "generate_cooperation_info",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "right_content_method": "generate_cooperation_financial",
        "right_fallback_text": B_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 13.5,
            "height_cm": 15.0,
        },
        "right_paragraph_spacing_pt": 0,
    },
    4: {
        "title": "1.2 永續策略與行動計畫",
        "layout": "A",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "left_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "sustainability_strategy_table_component.py"),
            "class": "SustainabilityStrategyTableComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 2.5,
                "top_cm": 3.0,
                "width_cm": 27.0,
                "height_cm": 13.5,
            },
        },
    },
    5: {
        "title": "1.3 ESG 核心支柱",
        "layout": "A",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "left_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "esg_pillars_table_component.py"),
            "class": "ESGPillarsTableComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 2.5,
                "top_cm": 3.0,
                "width_cm": 14.0,
                "height_cm": 13.0,
            },
        },
        "right_content_method": "generate_esg_pillars",
        "right_fallback_text": B_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 13.5,
            "height_cm": 15.0,
        },
        "right_paragraph_spacing_pt": 0,
    },
    6: {
        "title": "1.4 我們的董事會",
        "layout": "A",
        "content_method": "generate_governance_overview",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "right_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "board_skill_table_component.py"),
            "class": "BoardSkillTableComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 18.0,
                "top_cm": 3.0,
                "width_cm": 13.5,
                "height_cm": 13.0,
            },
        },
    },
    7: {
        "title": "1.5 ESG 路線圖（第一階段）",
        "layout": "A",
        "content_method": "generate_esg_roadmap_context",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "right_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "esg_roadmap_component.py"),
            "class": "ESGRoadmapComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "phase": "phase1",
                "left_cm": 19.0,
                "top_cm": 3.0,
                "width_cm": 13.5,
                "box_height_cm": 5.2,
                "gap_cm": 0.6,
            },
        },
    },
    8: {
        "title": "1.5 ESG 路線圖（第二階段）",
        "layout": "A",
        "content_method": "generate_esg_roadmap_context",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "right_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "esg_roadmap_component.py"),
            "class": "ESGRoadmapComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "phase": "phase2",
                "left_cm": 19.0,
                "top_cm": 3.0,
                "width_cm": 13.5,
                "box_height_cm": 5.2,
                "gap_cm": 0.6,
            },
        },
    },
    9: {
        "title": "2.1 利害關係人識別與分析",
        "layout": "A",
        "content_method": "generate_stakeholder_identify",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "right_content_method": "generate_stakeholder_analysis",
        "right_fallback_text": B_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 13.5,
            "height_cm": 15.0,
        },
        "right_paragraph_spacing_pt": 0,
    },
    10: {
        "title": "2.2 利害關係人群體",
        "layout": "A",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "left_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "stakeholder_group_table_component.py"),
            "class": "StakeholderGroupTableComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 3.0,
                "top_cm": 3.0,
                "width_cm": 27.0,
                "height_cm": 13.0,
            },
        },
    },
    11: {
        "title": "2.3 利害關係人關注與溝通",
        "layout": "B",
        "content_method": "generate_stakeholder_communication",
        "fallback_text": B_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "text_box": {
            "left_cm": 18.0,
            "top_cm": 3.2,
            "width_cm": 13.5,
            "height_cm": 14.0,
        },
        "left_media": {
            "mode": "single",
            "width_cm": 12.0,
            "area_left_cm": 3.0,
            "area_top_cm": 3.2,
        },
        # 暫時關閉第 11 頁的利害關係人溝通圖片，避免疑似問題圖片拖累整本 PPT 結構
        "image_paths": [],
    },
    12: {
        "title": "3.1 重大議題（敘述與視覺）",
        "layout": "A",
        "content_method": "generate_material_issues_text",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "image_path": os.path.join(ASSETS3_COMPANY_PATH, "preview_big_issues_bar_chart.png"),
        "image_width_cm": 13.5,
        "image_align": "right",
        "image_top_cm": 3.0,
        "image_left_cm": 18.0,
    },
    13: {
        "title": "3.2 重大性摘要",
        "layout": "A",
        "content_method": "generate_materiality_summary",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "image_path": os.path.join(ASSETS3_COMPANY_PATH, "preview_bubble_matrix.png"),
        "image_width_cm": 13.5,
        "image_align": "right",
        "image_top_cm": 3.0,
        "image_left_cm": 18.0,
    },
    14: {
        "title": "3.3 回應永續目標的產品與服務",
        "layout": "A",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "left_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "prod_sdg_table_component.py"),
            "class": "ProdSDGTableComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 2.5,
                "top_cm": 3.0,
                "width_cm": 27.0,
                "height_cm": 13.5,
            },
        },
    },
    15: {
        "title": "3.4 永續目標與 SDGs",
        "layout": "A",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "left_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "sdg_icon_grid.py"),
            "class": "SDGIconGrid",
            "method": "add_to_slide",
            "init_kwargs": {},
            "method_kwargs": {
                "icon_paths": SDG_ICON_PATHS,
                "left_cm": 3.0,
                "top_cm": 3.0,
                "cell_size_cm": 3.6,
                "gap_cm": 0.4,
                "columns": 3,
            },
        },
        "right_content_method": "generate_sdg_summary",
        "right_fallback_text": B_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 13.5,
            "height_cm": 15.0,
        },
        "right_paragraph_spacing_pt": 0,
    },
    16: {
        "title": "3.5 風險管理",
        "layout": "A",
        "content_method": "generate_risk_management_overview",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 11,
        "paragraph_spacing_pt": 0,
        "image_path": None,
        # 暫時遮蔽 3.5 的 flowchart component，避免觸發 PowerPoint 修復
        # "right_component": {
        #     "file": os.path.join(ASSETS3_COMPANY_PATH, "risk_flowchart_component.py"),
        #     "class": "RiskFlowchartComponent",
        #     "method": "add_to_slide",
        #     "init_kwargs": {
        #         "font_name": "Calibri",
        #     },
        #     "method_kwargs": {
        #         "left_cm": 19.0,
        #         "top_cm": 3.0,
        #         "width_cm": 13.5,
        #     },
        # },
    },
    17: {
        "title": "3.6 重大議題與風險管理",
        "layout": "A",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "left_component": {
            "file": os.path.join(ASSETS3_COMPANY_PATH, "risk_management_table_component.py"),
            "class": "RiskManagementTableComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 2.5,
                "top_cm": 3.0,
                "width_cm": 27.0,
                "height_cm": 13.0,
            },
        },
    },
}


