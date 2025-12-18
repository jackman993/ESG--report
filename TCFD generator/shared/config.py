"""共享配置"""
from pathlib import Path
import os

DESKTOP = Path(os.path.expanduser("~")) / "Desktop"
ESG_OUTPUT_ROOT = DESKTOP / "ESG_Output"

# 強制建立完整路徑
ESG_OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

OUTPUT_A_TCFD = ESG_OUTPUT_ROOT / "A_TCFD"
OUTPUT_A_TCFD.mkdir(parents=True, exist_ok=True)

OUTPUT_B_EMISSION = ESG_OUTPUT_ROOT / "B_Emission"
OUTPUT_B_EMISSION.mkdir(parents=True, exist_ok=True)

OUTPUT_D_COMPANY = ESG_OUTPUT_ROOT / "D_Company"
OUTPUT_D_COMPANY.mkdir(parents=True, exist_ok=True)

OUTPUT_C_ENVIRONMENT = ESG_OUTPUT_ROOT / "C_Environment"
OUTPUT_C_ENVIRONMENT.mkdir(parents=True, exist_ok=True)

OUTPUT_F_GOVSOCI = ESG_OUTPUT_ROOT / "F_Governance_Social"
OUTPUT_F_GOVSOCI.mkdir(parents=True, exist_ok=True)

# 後台資料夾
BACKEND_PATH = ESG_OUTPUT_ROOT / "_Backend"
BACKEND_PATH.mkdir(parents=True, exist_ok=True)
BACKEND_LOGS = BACKEND_PATH / "user_logs"
BACKEND_LOGS.mkdir(parents=True, exist_ok=True)

