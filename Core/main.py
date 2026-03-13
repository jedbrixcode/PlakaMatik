import os
import time
import sys
from data_processor import parse_input_data
from corel_engine import CorelAutomator
from print_handler import execute_print_merge_to_pdf

# System paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)

# Dynamic routing to the correct folders
INPUT_TXT_PATH = os.path.join(ROOT_DIR, "Csv", "temp_user_input.txt")
TEMPLATE_MV_PATH = os.path.join(ROOT_DIR, "CorelDRAW Templates", "MV_PLATE.cdr")
TEMPLATE_MC_PATH = os.path.join(ROOT_DIR, "CorelDRAW Templates", "Protocol Plates MC.cdr")
def run_pipeline():
    print("--- Starting LTO Automation Pipeline ---")
    
    # 1. Process the data
    data_records = parse_input_data(INPUT_TXT_PATH)
    
    if not data_records:
        print("Pipeline stopped: Data processing failed.")
        return

    # Determine type from the first record (or fallback to sys args/default)
    plate_type = "MV"
    if data_records and data_records[0].get("type"):
        pt = data_records[0]["type"].upper()
        if pt in ["MC", "MV"]:
            plate_type = pt
    # Also allow passing it as an arg just in case
    if len(sys.argv) > 1 and sys.argv[1].upper() in ["MV", "MC"]:
        plate_type = sys.argv[1].upper()

    print(f"Determined Plate Type: {plate_type}")
    template_path = TEMPLATE_MV_PATH if plate_type == "MV" else TEMPLATE_MC_PATH
    final_pdf_path = os.path.join(ROOT_DIR, f"LTO_Batch_Output_{plate_type}.pdf")

    # 2. Initialize the automation engine
    automator = CorelAutomator()
    automator.bypass_trial_screen()
    
    # 3. Connect and open template
    if automator.connect():
        automator.open_template(template_path)
        
        # Prevent locked file crashes by attempting to delete existing output
        if os.path.exists(final_pdf_path):
            try:
                os.remove(final_pdf_path)
                print("Deleted existing batch output file.")
            except Exception as e:
                print(f"Warning: Could not delete old PDF. Close it if it's open. {e}")

        # 4. Execute manual data merge
        time.sleep(2)

        merge_success = execute_print_merge_to_pdf(
            automator.corel,
            automator.doc,
            data_records,
            final_pdf_path,
            plate_type
        )

        if merge_success:
            print("--- pipeline completed successfully ---")
        else:
            print ("--- pipeline failed during manual merge ---")

if __name__ == "__main__":
    run_pipeline()
