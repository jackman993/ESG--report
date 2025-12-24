"""
ESG 環境篇報告生成器 - 配置檔案
"""
import os

# API 設定
ANTHROPIC_API_KEY = "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"
CLAUDE_MODEL = "claude-sonnet-4-20250514"  # 使用 Claude Sonnet 4（與 TCFD Generator 一致）

# 路徑設定（使用絕對路徑確保跨目錄調用正確）
import pathlib
ASSETS_PATH = str(pathlib.Path(__file__).parent / "assets")

# 環境篇基本配置
ENVIRONMENT_CONFIG = {
    'chapter_title': '4 環境永續',
    'total_pages': 17,
    'layout': 'A4_horizontal',
    'font_family': 'Microsoft JhengHei',
    'font_size': 12,  # 統一文字大小
    'table_split': '50_50'  # 左右各50%
}

# 每頁詳細配置
PAGE_CONFIGS = {
    'cover': {
        'layout': 'left_image_right_text',
        'title': '4 環境永續',
        'image': 'picture4-0Environment_cover.png',
        'content_method': 'generate_environmental_cover',
        'word_count': 250
    },
    'page_1': {
        'layout': 'left_image_right_text',
        'title': '4.1 環境政策與管理架構',
        'image': 'picture4-1sustainability_penal.png',
        'content_method': 'generate_sustainability_committee',
        'word_count': 250
        
    },
    'page_2': {
        'layout': 'left_text_right_image',
        'title': '4.2 環境政策四大面向',
        'image': 'picture4-2policy.png',
        'content_method': 'generate_policy_description',
        'word_count': 250

    },
    'page_3': {
        'layout': 'single_table',
        'title': 'TCFD 轉型風險',
        'table_file': 'tcfd_table_01_transformation_risk.py'
    },
    'page_4': {
        'layout': 'single_table',
        'title': 'TCFD 市場風險',
        'table_file': 'tcfd_table_02_market_risk.py'
    },
    'page_5': {
        'layout': 'single_table',
        'title': 'TCFD 實體風險',
        'table_file': 'tcfd_table_03_physical_risk.py'
    },
    'page_6': {
        'layout': 'single_table',
        'title': 'TCFD 溫升風險',
        'table_file': 'tcfd_table_04_temperature_rise.py'
    },
    'page_7': {
        'layout': 'single_table',
        'title': 'TCFD 效率機會',
        'table_file': 'tcfd_table_05_resource_efficiency.py'
    },
    'page_8': {
        'layout': 'single_table',
        'title': 'TCFD 主要架構說明',
        'table_file': 'TCFD_main.py'
    },
    'page_9': {
        'layout': 'left_text_right_image',
        'title': 'TCFD 風險矩陣分析',
        'image': 'picture4-3TCFD_matrix.png',
        'content_method': 'generate_tcfd_matrix_analysis',
        'word_count': 250
    },
    'page_10': {
        'layout': 'left_table_right_text',
        'title': '4.5 溫室氣體排放管理',
        'table_file': 'emission_table.py',
        'content_method': 'generate_ghg_calculation_method',
        'word_count': 250
    },
    'page_11': {
        'layout': 'left_image_right_text',
        'title': '電力使用與節能政策',
        'image': 'picture4-4ghg_pie_scope.png',
        'content_method': 'generate_electricity_policy',
        'word_count': 250
    },
    'page_12': {
        'layout': 'left_text_right_image',
        'title': '節能技術措施',
        'image': 'picture4-5ghg_bar.png',
        'content_method': 'generate_energy_efficiency_measures',
        'word_count': 250
    },
    'page_13': {
        'layout': 'left_image_right_text',
        'title': '4.6 綠色植育',
        'image': 'picture4-6plant.png',
        'content_method': 'generate_green_planting_program',
        'word_count': 250
    },
    'page_14': {
        'layout': 'left_text_right_image',
        'title': '4.7 水資源管理',
        'image': 'picture4-7water.png',
        'content_method': 'generate_water_management',
        'word_count': 250

    },
    'page_15': {
        'layout': 'left_text_right_image',
        'title': '4.8 廢棄物管理',
        'image': 'picture4-8waste.png',
        'content_method': 'generate_waste_management',
        'word_count': 250
    },
    'page_16': {
        'layout': 'left_text_right_image',
        'title': '4.9 環境教育與合作',
        'image': 'picture4-9ecowork.png',
        'content_method': 'generate_environmental_education',
        'word_count': 250
    }
}

# 頁面順序
PAGE_ORDER = ['cover', 'page_1', 'page_2', 'page_3', 'page_4', 'page_5', 'page_6', 
              'page_7', 'page_8', 'page_9', 'page_10', 'page_11', 'page_12', 
              'page_13', 'page_14', 'page_15', 'page_16']

# 圖片映射（所有圖片都在 assets 資料夾內）
ENVIRONMENT_IMAGE_MAPPING = {
    'cover': 'picture4-0Environment_cover.png',
    'sustainability_panel': 'picture4-1sustainability_penal.png',
    'policy': 'picture4-2policy.png',
    'tcfd_matrix': 'picture4-3TCFD_matrix.png',
    'ghg_pie': 'picture4-4ghg_pie_scope.png',
    'ghg_bar': 'picture4-5ghg_bar.png',
    'plant': 'picture4-6plant.png',
    'water': 'picture4-7water.png',
    'waste': 'picture4-8waste.png',
    'ecowork': 'picture4-9ecowork.png'
}

# TCFD 表格映射（所有 .py 檔案都在 assets 資料夾內）
TCFD_TABLES = {
    'transformation_risk': 'tcfd_table_01_transformation_risk.py',
    'market_risk': 'tcfd_table_02_market_risk.py',
    'physical_risk': 'tcfd_table_03_physical_risk.py',
    'temperature_rise': 'tcfd_table_04_temperature_rise.py',
    'resource_efficiency': 'tcfd_table_05_resource_efficiency.py',
    'main_framework': 'TCFD_main.py'
}