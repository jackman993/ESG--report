"""
Step 1 +1 步驟追蹤器
追蹤 +1 步驟（生成 150 字產業別分析）是否執行
"""
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional


LOG_DIR = Path(__file__).parent
TRACKER_FILE = LOG_DIR / "step1_plus1_tracker.json"


def log_plus1_step(session_id: str, status: str, details: Dict[str, Any] = None) -> None:
    """
    記錄 +1 步驟執行情況
    
    Args:
        session_id: Session ID
        status: 執行狀態（"started", "success", "failed"）
        details: 詳細資訊
    """
    tracker_data = {
        "session_id": session_id,
        "timestamp": datetime.now().isoformat(),
        "status": status,
        "details": details or {}
    }
    
    # 讀取現有記錄
    all_records = []
    if TRACKER_FILE.exists():
        try:
            with open(TRACKER_FILE, "r", encoding="utf-8") as f:
                all_records = json.load(f)
        except:
            all_records = []
    
    # 添加新記錄
    all_records.append(tracker_data)
    
    # 只保留最近 100 條記錄
    if len(all_records) > 100:
        all_records = all_records[-100:]
    
    # 寫入文件
    try:
        with open(TRACKER_FILE, "w", encoding="utf-8") as f:
            json.dump(all_records, f, ensure_ascii=False, indent=2)
        print(f"[+1 Tracker] 記錄 {status}: session_id={session_id}")
    except Exception as e:
        print(f"[+1 Tracker] 寫入失敗: {e}")


def check_plus1_execution(session_id: str) -> Optional[Dict[str, Any]]:
    """
    檢查 +1 步驟是否執行
    
    Args:
        session_id: Session ID
    
    Returns:
        執行記錄，如果沒有則返回 None
    """
    if not TRACKER_FILE.exists():
        return None
    
    try:
        with open(TRACKER_FILE, "r", encoding="utf-8") as f:
            all_records = json.load(f)
        
        # 查找該 session_id 的記錄
        for record in reversed(all_records):  # 從最新開始查找
            if record.get("session_id") == session_id:
                return record
    except:
        pass
    
    return None


def get_latest_plus1_status() -> Optional[Dict[str, Any]]:
    """
    獲取最新的 +1 步驟執行狀態
    
    Returns:
        最新的執行記錄
    """
    if not TRACKER_FILE.exists():
        return None
    
    try:
        with open(TRACKER_FILE, "r", encoding="utf-8") as f:
            all_records = json.load(f)
        
        if all_records:
            return all_records[-1]  # 返回最新的記錄
    except:
        pass
    
    return None


if __name__ == "__main__":
    # 測試
    test_session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_plus1_step(test_session_id, "started", {"test": True})
    log_plus1_step(test_session_id, "success", {"analysis_length": 150})
    
    # 檢查
    record = check_plus1_execution(test_session_id)
    print(f"檢查結果: {record}")
    
    latest = get_latest_plus1_status()
    print(f"最新狀態: {latest}")

