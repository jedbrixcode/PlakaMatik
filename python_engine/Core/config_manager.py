import os
import json
import config

def load_config(override_path=None):
    """
    Loads dynamic settings injected by the Flutter Frontend JSON Bridge.
    Falls back to factory defaults gracefully if bridging fails.
    """
    config_path = override_path if override_path else os.path.join(config.PLAKAMATIK_DIR, "config.json")
    
    # Factory defaults
    defaults = {
        "PRINTER_NAME": "Microsoft Print to PDF",
        "CORELDRAW_VISIBLE": False,
        "GLOBAL_OFFSETS": {
            "dx": 0.0,
            "dy": 0.0
        }
    }
    
    try:
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                user_config = json.load(f)
                
            # Merge dictionary allowing fallback for missing keys
            for k, v in user_config.items():
                if isinstance(v, dict) and k in defaults:
                    defaults[k].update(v)
                else:
                    defaults[k] = v
                    
            print(f"JSON Bridge connected successfully. Loaded presets.")
        else:
            print(f"JSON Bridge skipped: config.json not found. Operating on factory defaults.")
    except Exception as e:
        print(f"JSON Bridge error: {e}. Falling back to safe factory defaults.")
        
    return defaults
