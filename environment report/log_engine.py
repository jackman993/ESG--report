"""
Log 引擎 - 記錄 TCFD 和 Emission 結果供其他章節使用
"""
import json
import os
from pathlib import Path
from datetime import datetime

# Log 輸出目錄
LOG_DIR = Path(__file__).parent / "output" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)


class LogEngine:
    """Log 引擎 - 記錄並提供跨章節資料"""
    
    def __init__(self):
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.logs = {
            "timestamp": datetime.now().isoformat(),
            "tcfd": {},
            "emission": {},
            "prompts": {}
        }
    
    # ============ TCFD Log ============
    
    def log_tcfd_transformation(self, industry, key_points):
        """
        記錄 TCFD 轉型風險
        用於：CEO的話、治理段(法規與政策)
        """
        self.logs["tcfd"]["transformation"] = {
            "industry": industry,
            "key_points": key_points,
            "usage": ["CEO的話", "治理段-法規與政策"]
        }
        
        # 生成 prompt 上下文
        policy_text = "；".join(key_points[:3]) if key_points else "碳費徵收、環評法規"
        self.logs["prompts"]["ceo_policy_context"] = f"根據TCFD轉型風險分析，{industry}面臨的政策法規挑戰包括：{policy_text}。"
        self.logs["prompts"]["governance_policy_context"] = f"針對{industry}的法規環境，主要政策包括：{policy_text}。"
        
        print(f"  ✓ Log: TCFD 轉型風險 (用於 CEO的話、治理段)")
    
    def log_tcfd_market(self, industry, key_points):
        """
        記錄 TCFD 市場風險
        用於：關係人分析段
        """
        self.logs["tcfd"]["market"] = {
            "industry": industry,
            "key_points": key_points,
            "usage": ["關係人分析段"]
        }
        
        # 生成 prompt 上下文
        market_text = "；".join(key_points[:3]) if key_points else "消費者偏好變化、B2B客戶ESG要求"
        self.logs["prompts"]["stakeholder_context"] = f"根據TCFD市場風險分析，{industry}的主要市場趨勢包括：{market_text}。"
        
        print(f"  ✓ Log: TCFD 市場風險 (用於 關係人分析段)")
    
    # ============ Emission Log ============
    
    def log_emission(self, emission_data):
        """
        記錄 Emission 結果
        用於：社會責任段、節能技術段
        """
        total = emission_data.get("total", 153.45)
        scope1 = emission_data.get("scope1", {}).get("subtotal", 5.34)
        scope2 = emission_data.get("scope2", {}).get("subtotal", 148.11)
        scope3 = emission_data.get("scope3", {}).get("subtotal", 0.00)
        
        main_percent = round(scope2 / total * 100, 2) if total > 0 else 0
        
        self.logs["emission"] = {
            "data_year": emission_data.get("data_year", "2024"),
            "total": total,
            "scope1": scope1,
            "scope2": scope2,
            "scope3": scope3,
            "main_source": "外購電力",
            "main_source_percent": main_percent,
            "usage": ["社會責任段", "節能技術段"]
        }
        
        # 生成 prompt 上下文
        self.logs["prompts"]["social_emission_context"] = f"本公司{emission_data.get('data_year', '2024')}年溫室氣體排放總量為{total} tCO₂e，積極落實減碳目標。"
        self.logs["prompts"]["energy_efficiency_context"] = f"根據碳盤查結果，外購電力(範疇二)佔總排放{main_percent}%，為主要減碳重點。我們採取以下節能措施："
        
        print(f"  ✓ Log: Emission 結果 (用於 社會責任段、節能技術段)")
    
    # ============ 取得 Prompt 上下文 ============
    
    def get_prompt_context(self, context_type):
        """
        取得特定類型的 prompt 上下文
        
        context_type:
        - ceo_policy_context: CEO的話用
        - governance_policy_context: 治理段用
        - stakeholder_context: 關係人分析用
        - social_emission_context: 社會責任段用
        - energy_efficiency_context: 節能技術段用
        """
        return self.logs["prompts"].get(context_type, "")
    
    def get_all_prompts(self):
        """取得所有 prompt 上下文"""
        return self.logs["prompts"]
    
    # ============ 儲存與載入 ============
    
    def save(self, filename=None):
        """儲存 Log 到檔案"""
        if filename is None:
            filename = LOG_DIR / f"report_log_{self.timestamp}.json"
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.logs, f, ensure_ascii=False, indent=2)
        
        print(f"  ✓ Log 已儲存: {filename}")
        return filename
    
    @classmethod
    def load_latest(cls):
        """載入最新的 Log 檔案"""
        log_files = list(LOG_DIR.glob("report_log_*.json"))
        if not log_files:
            print("  ⚠ 未找到 Log 檔案")
            return None
        
        latest_file = max(log_files, key=os.path.getmtime)
        
        with open(latest_file, 'r', encoding='utf-8') as f:
            logs = json.load(f)
        
        engine = cls()
        engine.logs = logs
        print(f"  ✓ 載入 Log: {latest_file}")
        return engine


# ============ 測試 ============
if __name__ == "__main__":
    log = LogEngine()
    
    # 模擬 TCFD 資料
    log.log_tcfd_transformation("鋁門窗業", [
        "碳費徵收影響成本結構約5-10%",
        "環評法規趨嚴，需投資污染防治設備",
        "低碳鋁材認證需求增加"
    ])
    
    log.log_tcfd_market("鋁門窗業", [
        "消費者偏好綠建材與節能門窗",
        "B2B客戶要求供應商提供碳足跡資料",
        "ESG評比影響企業合作機會"
    ])
    
    # 模擬 Emission 資料
    log.log_emission({
        "data_year": "2024",
        "total": 153.45,
        "scope1": {"subtotal": 5.34},
        "scope2": {"subtotal": 148.11},
        "scope3": {"subtotal": 0.00}
    })
    
    # 顯示所有 prompt 上下文
    print("\n[Prompt 上下文]")
    for key, value in log.get_all_prompts().items():
        print(f"  {key}:")
        print(f"    {value}\n")
    
    # 儲存
    log.save()

