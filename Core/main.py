import os
from data_processor import export_data_to_rtf
from corel_engine import CorelAutomator

# System paths
# BASE_DIR is the 'Core' folder
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ROOT_DIR is 'PLATE Manufaturing Layout maker'
ROOT_DIR = os.path.dirname(BASE_DIR)

# Dynamic routing to the correct folders
INPUT_TXT_PATH = os.path.join(ROOT_DIR, "Csv", "temp_user_input.txt")
OUTPUT_RTF_PATH = os.path.join(ROOT_DIR, "Csv", "output_data.rtf")
TEMPLATE_PATH = os.path.join(ROOT_DIR, "CorelDRAW Templates", "MV_PLATE.cdr")

def run_pipeline():
    print("--- Starting LTO Automation Pipeline ---")
    
    # 1. Process the data
    data_success = export_data_to_rtf(INPUT_TXT_PATH, OUTPUT_RTF_PATH)
    
    if not data_success:
        print("Pipeline stopped: Data processing failed.")
        return

    # 2. Initialize the automation engine
    automator = CorelAutomator()
    automator.bypass_trial_screen()
    
    # 3. Connect and open template
    if automator.connect():
        automator.open_template(TEMPLATE_PATH)
        
        # Next phase placeholder: Bind RTF and print
        # automator.bind_and_print(OUTPUT_RTF_PATH)

if __name__ == "__main__":
    run_pipeline()