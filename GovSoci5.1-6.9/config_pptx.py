"""
PPT configuration for governance + social sections (A/B templates only).
"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# API settings
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
# 使用 handdrawppt.pptx 作為母版模板（A4 landscape）
SEED_TEMPLATE_PATH = str(BASE_DIR.parent / "assets2" / "handdrawppt.pptx")
OUTPUT_PATH = str(BASE_DIR / "output")
# Assets 已移到上一層目錄
ASSETS2_PATH = str(BASE_DIR.parent / "assets2")
ASSETS3_PATH = str(BASE_DIR.parent / "assets3")

# Core PPT settings
PPT_CONFIG = {
    "seed_template": SEED_TEMPLATE_PATH,
    "output_filename": "ESG_PPT_AB_7slides_cleanup.pptx",
    "total_slides": 18,  # 增加2張章節封面
    "font_family": "Calibri",
    "font_size": 12,
    "watermark_text": "Sustainability Report",
}

# Common text snippets (will be replaced by LLM later, but kept as fallback)
A_TEXT_FALLBACK = (
    "This page uses Layout A, which positions narrative copy in the left column and places a hero image on the right. "
    "The text area remains constrained within its textbox, ensuring the photograph or chart on the right never overlaps the narrative." )

B_TEXT_FALLBACK = (
    "This page demonstrates Layout B. Supporting imagery is stacked inside the left media frame while the descriptive copy sits on the right. "
    "Each image stays within its bounds so it never collides with the text column." )

SLIDE_CONFIGS = {
    # 治理段封面
    1: {
        "title": "",
        "layout": "cover",
        "paragraphs": 0,
        "image_path": os.path.join(ASSETS2_PATH, "5-0cover.png"),
        "image_width_cm": 29.7,  # A4 寬度
        "image_height_cm": 16.0,
        "image_left_cm": 0,
        "image_top_cm": 0,
        "cover_text": {
            "left_cm": 3.0,
            "top_cm": 6.0,
            "lines": [
                {"text": "治理", "font_size": 45, "bold": True},
                {"text": "誠信塑造我們的未來", "font_size": 32, "bold": False},
            ],
            "font_color": (128, 128, 128),  # Gray
        },
    },
    # 治理段內容
    2: {
        "title": "5.1 治理概覽",
        "layout": "A",
        "content_method": "generate_governance_overview",  # 啟用 LLM 文字生成
        "fallback_text": "",
        "image_path": os.path.join(ASSETS2_PATH, "5.1gover_structure.png"),
        "image_width_cm": 14.0,
        "image_align": "right",
        "paragraphs": 2,  # 啟用文字框
        "text_font_size": 12,
        "title_font_pt": 20,
        "title_top_cm": 1.2,
        "title_left_cm": 2.5,
        "right_notes": [
            {"text": "Female directors: XX% (pending)", "color": (198, 38, 38)},
            {"text": "Independent directors: XX% (pending)", "color": (198, 38, 38)},
            {"text": "Minority directors: XX% (pending)", "color": (198, 38, 38)},
        ],
    },
    3: {
        "title": "5.2 性別平等",
        "layout": "B",
        "content_method": "generate_gender_equality_overview",  # 啟用 LLM 文字生成
        "fallback_text": "",
        "image_paths": [
            os.path.join(ASSETS2_PATH, "5.2.1gender.png"),
            os.path.join(ASSETS2_PATH, "5.2.2gender2.png"),
        ],
        "left_media": {
            "mode": "stack",
            "width_cm": 8.0,
            "gap_cm": 3.0,
            "area_left_cm": 2.9,
            "area_top_cm": 2.8,
            "area_height_cm": 13.0,
            "overlay_box": {
                "left_cm": 2.7,
                "top_cm": 2.7,
                "width_cm": 9.5,
                "height_cm": 13.5,
            },
        },
        "paragraphs": 2,  # 啟用文字框
        "text_font_size": 12,
        "title_font_pt": 20,
        "title_top_cm": 1.2,
        "title_left_cm": 2.5,
    },
    4: {
        "title": "5.3 法規對齊概覽",
        "layout": "A",
        "content_method": "generate_legal_alignment_overview",  # 啟用 LLM 文字生成
        "fallback_text": "",
        "paragraphs": 2,  # 啟用文字框
        "text_font_size": 12,
        "image_path": os.path.join(ASSETS2_PATH, "5.3legal.png"),
        "image_width_cm": 14.0,
        "image_align": "right",
        "image_top_cm": 2.8,
        "title_font_pt": 20,
        "title_top_cm": 1.2,
        "title_left_cm": 2.5,
    },
    5: {
        "title": "5.4 法律遵循",
        "layout": "A",
        "content_method": "generate_legal_appliance_overview",  # 啟用 LLM 文字生成
        "fallback_text": "",
        "paragraphs": 2,  # 啟用文字框
        "text_font_size": 12,
        "title_font_pt": 20,
        "title_top_cm": 1.2,
        "title_left_cm": 2.5,
        # 測試：啟用簡單表格 5.4
        "image_path": None,
        "right_component": {
            "file": os.path.join(ASSETS2_PATH, '5.4legal_applaince.py'),
            "class": 'EuropeanComplianceTable',
            "method": 'add_to_slide',
            "init_kwargs": {
                "font_name": "Arial",
                "colors": None,
            },
            "method_kwargs": {
                "title": "",
                "show_title": False,
                "board_left_in": 7.6,
                "board_top_in": 1.5,
                "board_width_in": 5.4,
                "board_height_in": 3.8,
                "table_font_size": 10,
            },
        },
        "left_notes": [
            {"text": "Base pay: XX% (pending)", "color": (198, 38, 38)},
            {"text": "STI: XX% (pending)", "color": (198, 38, 38)},
            {"text": "LTI: XX% (pending)", "color": (198, 38, 38)},
            {"text": "ESG weighting: XX% (pending)", "color": (198, 38, 38)},
        ],
    },
    6: {
        "title": "5.5 稽核委員會季度計畫",
        "layout": "B",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "title_font_pt": 20,
        "title_top_cm": 1.2,
        "title_left_cm": 2.5,
        # 恢復 5.5 component 引擎（正常）
        "left_media": {
            "mode": "component",
            "file": os.path.join(ASSETS2_PATH, "5.5audit_board.py"),
            "class": "AuditBoardComponent",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Arial",
                "title": "",
            },
            "method_kwargs": {
                "title": "",
                "panel_left_in": 1.1,
                "panel_top_in": 1.4,
                "card_spacing_in": 0.25,
            },
        },
    },
    7: {
        "title": "5.6 稽核委員會概覽",
        "layout": "C",
        "content_method": "generate_supervisory_board_overview",  # 使用 LLM 生成中文文本
        "fallback_text": "",
        "text_font_size": 12,
        "title_font_pt": 20,
        "title_top_cm": 1.2,
        "title_left_cm": 2.5,
        "top_paragraphs": 1,
        "bottom_paragraphs": 1,
        "top_box_left_cm": 2.2,
        "top_box_top_cm": 3.0,
        "top_box_width_cm": 27.0,
        "top_box_height_cm": 5.0,
        "bottom_box_left_cm": 2.2,
        "bottom_box_width_cm": 27.0,
        "bottom_box_height_cm": 4.5,
        # 移除硬編碼的英文文本，改用 content_method 生成的中文文本
        # "top_text": ... (已移除)
        # "bottom_text": ... (已移除)
        # 恢復 5.6 component 引擎（正常）
        "center_media": {
            "mode": "component",
            "file": os.path.join(ASSETS2_PATH, '5.6supervisory_board_chart.py'),
            "class": 'SupervisoryBoardChart',
            "method": 'add_to_slide',
            "init_kwargs": {
                "scale": 0.78,
                "font_scale": 0.72,
            },
            "method_kwargs": {
                "x": 1.2,
                "y": 2.4,
            },
            "top_cm": 8.0,
            "bottom_text_top_cm": 15.0,
        },
    },
    8: {
        "title": "5.7 董事薪酬架構",
        "layout": "B",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "title_font_pt": 20,
        "title_top_cm": 1.2,
        "title_left_cm": 2.5,
        # 測試：啟用 5.7 component 引擎
        "left_media": {
            "mode": "component",
            "file": os.path.join(ASSETS2_PATH, "5.7remuneration.py"),
            "class": "BoardIncentiveFramework",
            "method": "add_to_slide",
            "init_kwargs": {
                "scale": 1.05,
            },
            "method_kwargs": {
                "x": 0.8,
                "y": 1.2,
            },
        },
    },
    # 社會段封面
    9: {
        "title": "",
        "layout": "cover",
        "paragraphs": 0,
        "image_path": os.path.join(ASSETS3_PATH, "6-0cover.png"),
        "image_width_cm": 29.7,  # A4 寬度
        "image_height_cm": 16.0,
        "image_left_cm": 0,
        "image_top_cm": 0,
        "cover_text": {
            "left_cm": 3.0,
            "top_cm": 6.0,
            "lines": [
                {"text": "Social", "font_size": 45, "bold": True},
                {"text": "Empowering People", "font_size": 32, "bold": False},
                {"text": "Inspiring Change", "font_size": 32, "bold": False},
            ],
            "font_color": (255, 255, 255),  # White
        },
    },
    # 社會段內容
    10: {
        "title": "6.1 社區參與與社會投資",
        "layout": "A",
        "content_method": "generate_social_community_investment",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        # 測試：啟用簡單表格 6.1
        "right_component": {
            "file": os.path.join(ASSETS3_PATH, "6.1hr_overview.py"),
            "class": "CommunityInvestmentTable",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 18.0,
                "top_cm": 3.0,
                "width_cm": 14.0,
            },
        },
        "left_notes": [
            {"text": "Community investment: $X Million (pending)", "color": (198, 38, 38)},
            {"text": "Volunteer hours: XX,XXX (pending)", "color": (198, 38, 38)},
        ],
    },
    11: {
        "title": "6.2 員工健康、安全與福祉",
        "layout": "A",
        "content_method": "generate_social_health_safety",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        "image_path": os.path.join(ASSETS3_PATH, "6.2health.png"),
        "image_width_cm": 14.0,
        "image_align": "right",
        "image_top_cm": 3.0,
        "image_left_cm": 18.0,
    },
    12: {
        "title": "6.3 多元、包容與平等機會",
        "layout": "A",
        "content_method": "generate_social_diversity_policies",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        "right_content_method": "generate_social_diversity_kpis",
        "right_fallback_text": B_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 13.5,
            "height_cm": 15.0,
        },
        "right_paragraph_spacing_pt": 0,
        "right_notes": [
            {"text": "Gender pay gap: XX% (pending)", "color": (198, 38, 38)},
        ],
    },
    13: {
        "title": "6.4 勞動權利與公平就業實務",
        "layout": "A",
        "content_method": "generate_social_labor_rights",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        "right_content_method": "generate_social_fair_employment",
        "right_fallback_text": B_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 13.5,
            "height_cm": 15.0,
        },
        "right_paragraph_spacing_pt": 0,
        "right_notes": [
            {"text": "Community investment: $X Million (pending)", "color": (198, 38, 38)},
            {"text": "Volunteer hours: XX,XXX (pending)", "color": (198, 38, 38)},
        ],
    },
    14: {
        "title": "6.5 社會行動計畫",
        "layout": "A",
        "content_method": None,
        "fallback_text": "",
        "paragraphs": 0,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        # 測試：啟用簡單表格 6.5
        "left_component": {
            "file": os.path.join(ASSETS3_PATH, "6.5social_action_table.py"),
            "class": "SocialActionPlanTable",
            "method": "add_to_slide",
            "init_kwargs": {
                "font_name": "Calibri",
            },
            "method_kwargs": {
                "left_cm": 2.3,
                "top_cm": 3.0,
                "width_cm": 14.0,
            },
        },
        "right_content_method": "generate_social_action_plan_overview",
        "right_fallback_text": A_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 13.5,
            "height_cm": 15.0,
        },
        "right_paragraph_spacing_pt": 0,
    },
    15: {
        "title": "6.6 視覺展示：社會影響實踐",
        "layout": "B",
        "content_method": "generate_social_showcase_intro",
        "fallback_text": B_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        "text_box": {
            "left_cm": 3.0,
            "top_cm": 3.2,
            "width_cm": 10.0,
            "height_cm": 14.0,
        },
        "image_paths": [
            os.path.join(ASSETS3_PATH, "6.6.1youth.png"),
            os.path.join(ASSETS3_PATH, "6.6.2plant.png"),
            os.path.join(ASSETS3_PATH, "6.6.3animal.png"),
        ],
        "captions": [
            "Empowering youth through learning",
            "Building awareness for sustainability",
            "Caring for animals with compassion",
        ],
        "left_media": {
            "mode": "gallery",
            "area_left_cm": 14.5,
            "area_top_cm": 3.0,
             "area_height_cm": 14.0,
            "width_cm": 17.0,
            "gap_cm": 0.5,
            "caption_font_size": 10,
            "widths_cm": [7.0, 7.0, 9.0],
            "layout": "left_stack_right_single",
            "column_gap_cm": 1.0,
        },
    },
    16: {
        "title": "6.7 社會影響流程",
        "layout": "B",
        "content_method": "generate_social_flow_explanation",
        "fallback_text": B_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        # 6.7 使用 PNG 圖片（SocialImpactFlow 組件已刪除，改用圖片替代）
        "image_paths": [
            os.path.join(ASSETS3_PATH, "6.7社會影響.png"),
        ],
        "left_media": {
            "mode": "single",
            "area_left_cm": 2.9,
            "area_top_cm": 3.2,
            "width_cm": 11.0,
        },
    },
    17: {
        "title": "6.8 產品責任與客戶福祉",
        "layout": "A",
        "content_method": "generate_social_product_responsibility",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        "right_content_method": "generate_social_customer_welfare",
        "right_fallback_text": B_TEXT_FALLBACK,
        "right_paragraphs": 2,
        "right_text_box": {
            "left_cm": 18.4,
            "top_cm": 3.2,
            "width_cm": 12.8,
            "height_cm": 12.8,
        },
        "right_paragraph_spacing_pt": 0,
        "right_text_color_rgb": (255, 255, 255),
        "right_background": {
            "image_path": os.path.join(ASSETS3_PATH, "6.8right.png"),
            "left_cm": 18.0,
            "top_cm": 3.0,
            "width_cm": 14.0,
            "height_cm": 13.0,
        },
    },
    18: {
        "title": "6.9 社會創新與包容性經濟參與",
        "layout": "A",
        "content_method": "generate_social_innovation",
        "fallback_text": A_TEXT_FALLBACK,
        "paragraphs": 2,
        "text_font_size": 12,
        "paragraph_spacing_pt": 0,
        "right_content_method": "generate_social_inclusive_economy",
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
}

