"""
訂閱管理器：管理用戶配額和使用統計
"""
import json
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Optional, Tuple

class SubscriptionManager:
    """管理用戶訂閱配額和使用統計"""
    
    def __init__(self, db_path: Path):
        """
        初始化訂閱管理器
        
        Args:
            db_path: 訂閱資料庫文件路徑
        """
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._load_db()
    
    def _load_db(self):
        """載入訂閱資料庫"""
        if self.db_path.exists():
            try:
                with open(self.db_path, 'r', encoding='utf-8') as f:
                    self.db = json.load(f)
            except Exception as e:
                print(f"[WARN] 無法載入訂閱資料庫: {e}，使用預設值")
                self._init_db()
        else:
            self._init_db()
    
    def _init_db(self):
        """初始化資料庫"""
        self.db = {
            "users": {},  # {user_id: {plan, start_date, requests_today, requests_month, last_reset, last_request_date}}
            "settings": {
                "max_requests_per_day": 100,
                "max_requests_per_month": 2000,
                "free_plan_daily": 10,
                "free_plan_monthly": 100
            }
        }
        self._save_db()
    
    def _save_db(self):
        """保存訂閱資料庫"""
        try:
            with open(self.db_path, 'w', encoding='utf-8') as f:
                json.dump(self.db, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[ERROR] 無法保存訂閱資料庫: {e}")
    
    def get_user_id(self, session_id: str, subscription_code: Optional[str] = None) -> str:
        """
        獲取或創建用戶 ID
        
        Args:
            session_id: Streamlit session ID
            subscription_code: 用戶輸入的訂閱碼（可選）
        
        Returns:
            用戶 ID
        """
        # 優先使用訂閱碼
        if subscription_code and subscription_code.strip():
            return f"sub_{subscription_code.strip()}"
        
        # 否則使用 session ID
        return f"session_{session_id}"
    
    def check_quota(self, user_id: str) -> Tuple[bool, str]:
        """
        檢查用戶配額
        
        Args:
            user_id: 用戶 ID
        
        Returns:
            (是否可以繼續, 訊息)
        """
        today = datetime.now().date().isoformat()
        
        # 新用戶，創建記錄
        if user_id not in self.db["users"]:
            self.db["users"][user_id] = {
                "plan": "free",
                "start_date": today,
                "requests_today": 0,
                "requests_month": 0,
                "last_reset": today,
                "last_request_date": today
            }
            self._save_db()
        
        user = self.db["users"][user_id]
        
        # 檢查是否需要重置每日配額
        if user["last_reset"] != today:
            user["requests_today"] = 0
            user["last_reset"] = today
            self._save_db()
        
        # 檢查是否需要重置每月配額（每月 1 號重置）
        current_date = datetime.now().date()
        last_reset_date = datetime.fromisoformat(user["last_reset"]).date()
        if current_date.month != last_reset_date.month or current_date.year != last_reset_date.year:
            user["requests_month"] = 0
            user["last_reset"] = today
            self._save_db()
        
        # 根據方案獲取配額限制
        if user["plan"] == "free":
            max_daily = self.db["settings"]["free_plan_daily"]
            max_monthly = self.db["settings"]["free_plan_monthly"]
        else:
            max_daily = self.db["settings"]["max_requests_per_day"]
            max_monthly = self.db["settings"]["max_requests_per_month"]
        
        # 檢查每日配額
        if user["requests_today"] >= max_daily:
            return False, f"今日配額已用完（{max_daily} 次）。如需增加配額，請聯繫管理員升級訂閱。"
        
        # 檢查每月配額
        if user["requests_month"] >= max_monthly:
            return False, f"本月配額已用完（{max_monthly} 次）。如需增加配額，請聯繫管理員升級訂閱。"
        
        return True, "OK"
    
    def record_request(self, user_id: str):
        """
        記錄一次請求
        
        Args:
            user_id: 用戶 ID
        """
        today = datetime.now().date().isoformat()
        
        if user_id not in self.db["users"]:
            self.db["users"][user_id] = {
                "plan": "free",
                "start_date": today,
                "requests_today": 0,
                "requests_month": 0,
                "last_reset": today,
                "last_request_date": today
            }
        
        user = self.db["users"][user_id]
        
        # 重置每日配額（如果需要）
        if user["last_reset"] != today:
            user["requests_today"] = 0
            user["last_reset"] = today
        
        # 重置每月配額（如果需要）
        current_date = datetime.now().date()
        last_reset_date = datetime.fromisoformat(user["last_reset"]).date()
        if current_date.month != last_reset_date.month or current_date.year != last_reset_date.year:
            user["requests_month"] = 0
            user["last_reset"] = today
        
        user["requests_today"] += 1
        user["requests_month"] += 1
        user["last_request_date"] = today
        self._save_db()
    
    def get_usage_stats(self, user_id: str) -> Dict:
        """
        獲取用戶使用統計
        
        Args:
            user_id: 用戶 ID
        
        Returns:
            使用統計字典
        """
        if user_id not in self.db["users"]:
            return {
                "plan": "free",
                "requests_today": 0,
                "requests_month": 0,
                "max_daily": self.db["settings"]["free_plan_daily"],
                "max_monthly": self.db["settings"]["free_plan_monthly"]
            }
        
        user = self.db["users"][user_id]
        
        # 根據方案獲取配額限制
        if user["plan"] == "free":
            max_daily = self.db["settings"]["free_plan_daily"]
            max_monthly = self.db["settings"]["free_plan_monthly"]
        else:
            max_daily = self.db["settings"]["max_requests_per_day"]
            max_monthly = self.db["settings"]["max_requests_per_month"]
        
        return {
            "plan": user["plan"],
            "requests_today": user["requests_today"],
            "requests_month": user["requests_month"],
            "max_daily": max_daily,
            "max_monthly": max_monthly
        }
    
    def upgrade_plan(self, user_id: str, plan: str = "premium"):
        """
        升級用戶方案（管理員功能）
        
        Args:
            user_id: 用戶 ID
            plan: 方案名稱
        """
        if user_id not in self.db["users"]:
            self.db["users"][user_id] = {
                "plan": plan,
                "start_date": datetime.now().date().isoformat(),
                "requests_today": 0,
                "requests_month": 0,
                "last_reset": datetime.now().date().isoformat(),
                "last_request_date": datetime.now().date().isoformat()
            }
        else:
            self.db["users"][user_id]["plan"] = plan
        
        self._save_db()
    
    def get_all_users(self) -> Dict:
        """獲取所有用戶統計（管理員功能）"""
        return self.db["users"]

