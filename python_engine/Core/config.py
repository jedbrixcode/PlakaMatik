import os
import sys
from datetime import datetime

# System paths
if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
    ROOT_DIR = os.path.dirname(EXE_DIR)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ROOT_DIR = os.path.dirname(BASE_DIR)

FLUTTER_ROOT_DIR = os.path.dirname(ROOT_DIR)

# Dynamic routing to the correct Windows profile folders
PLAKAMATIK_DIR = os.path.join(os.path.expanduser("~"), "Documents", "PlakaMatik Files")

# Input / Output paths
INPUT_TXT_PATH = os.path.join(PLAKAMATIK_DIR, "csv", "flutter_user_input.txt")
OUTPUTS_DIR = os.path.join(PLAKAMATIK_DIR, "Outputs")
TEMP_PREVIEWS_DIR = os.path.join(PLAKAMATIK_DIR, "temp_previews")

# Logging
SESSION_ID = datetime.now().strftime("%Y%m%d_%H%M")
LOGS_DIR = os.path.join(PLAKAMATIK_DIR, "Logs")

# Ensure environment directories exist
for directory in [LOGS_DIR, OUTPUTS_DIR, TEMP_PREVIEWS_DIR]:
    if not os.path.exists(directory):
        os.makedirs(directory)


def _find_template(filename):
    """
    Searches for a CDR template file by trying candidate paths in order.
    Falls back to a full recursive walk of PlakaMatik Files if not found.
    Returns the absolute path string if found, or None.
    """
    candidates = [
        # User put them directly in 'CorelDRAW Templates/Main Templates'
        os.path.join(PLAKAMATIK_DIR, "CorelDRAW Templates", "Main Templates", filename),
        # User put them directly in 'CorelDRAW Templates'
        os.path.join(PLAKAMATIK_DIR, "CorelDRAW Templates", filename),
        # User put them directly in PlakaMatik Files root
        os.path.join(PLAKAMATIK_DIR, filename),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c

    # Last resort: walk the entire PlakaMatik Files tree
    for root, dirs, files in os.walk(PLAKAMATIK_DIR):
        if filename in files:
            return os.path.join(root, filename)

    return None


# Resolve template paths at startup and log clearly
_mv_resolved = _find_template("MV_PLATE.cdr")
_mc_resolved = _find_template("MC_PLATE.cdr")

TEMPLATE_MV_PATH = _mv_resolved or os.path.join(PLAKAMATIK_DIR, "CorelDRAW Templates", "Main Templates", "MV_PLATE.cdr")
TEMPLATE_MC_PATH = _mc_resolved or os.path.join(PLAKAMATIK_DIR, "CorelDRAW Templates", "Main Templates", "MC_PLATE.cdr")

PLATE_TEMPLATE_DIR = os.path.dirname(TEMPLATE_MV_PATH)
