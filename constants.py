# --- FILE: app/constants.py ---

# Centralized constants for DataFrame column names and other recurring strings.

class Columns:
    """Holds all DataFrame column names as constants."""
    # Invoice columns from source files
    INVOICE_NUMBER = "NÚMERO"
    ISSUE_DATE = "DATA EMISSÃO"
    VALUE = "VALOR"
    RATE = "ALÍQUOTA"
    SERVICE_DESCRIPTION = "DISCRIMINAÇÃO DOS SERVIÇOS"
    ACTIVITY_CODE = "CÓDIGO DA ATIVIDADE" # <<< FIX: Corrected typo from ATIVITY to ATIVIDADE
    
    # AI Analysis columns
    ORIGINAL_PROB = "original_prob"
    PREDICTED_CATEGORY = "predicted_category"
    PREDICTION_PROB = "prediction_prob"
    CATEGORY_IS_OK = "category_is_ok"
    LOCATION_ALERT = "location_alert"
    
    # Rules Analysis columns
    BROKEN_RULE_DETAILS = "broken_rule_details"
    CORRECT_RATE = "correct_rate"
    ACTIVITY_DESC = "activity_desc"
    STATUS_LEGAL = "status_legal"
    
    # Company/Cadastro columns
    CNPJ = "cnpj"
    RAZAO_SOCIAL = "razao_social"
    IMU = "imu"
    ENDERECO = "endereco"
    CEP = "cep"
    EPAF_NUMERO = "epaf_numero"

# Other constants
APP_NAME = "Caronte FFRM"
SESSION_FILE_PREFIX = "session_"
APP_VERSION = "1.1.0"  # Update this before every pyinstaller build
GITHUB_REPO_OWNER = "Ostrensky" 
GITHUB_REPO_NAME = "Caronte_FFRM"

GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
