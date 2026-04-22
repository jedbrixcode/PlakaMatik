import os
from datetime import datetime

# System paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
FLUTTER_ROOT_DIR = os.path.dirname(ROOT_DIR)

# Subfolder for Templates
PLATE_TEMPLATE_DIR = os.path.join(ROOT_DIR, "CorelDRAW Templates", "Main Templates")

# Dynamic routing to the correct Windows profile folders
PLAKAMATIK_DIR = os.path.join(os.path.expanduser("~"), "Documents", "PlakaMatik Files")
INPUT_TXT_PATH = os.path.join(PLAKAMATIK_DIR, "csv", "flutter_user_input.txt")
OUTPUTS_DIR = os.path.join(PLAKAMATIK_DIR, "Outputs")
TEMP_PREVIEWS_DIR = os.path.join(PLAKAMATIK_DIR, "temp_previews")

# Template Mapping
TEMPLATE_MV_PATH = os.path.join(PLATE_TEMPLATE_DIR, "MV_PLATE.cdr")
TEMPLATE_MC_PATH = os.path.join(PLATE_TEMPLATE_DIR, "MC_PLATE.cdr")

# Generate Session ID (yyyyMMdd_HHmm) and logging components
SESSION_ID = datetime.now().strftime("%Y%m%d_%H%M")
LOGS_DIR = os.path.join(PLAKAMATIK_DIR, "Logs")

# Ensure environment directories exist
for directory in [LOGS_DIR, OUTPUTS_DIR, TEMP_PREVIEWS_DIR]:
    if not os.path.exists(directory):
        os.makedirs(directory)
