import os
import time
import argparse

# Custom Modules
import config
from engine_logger import init_logger
from data_processor import parse_input_data
from corel_engine import CorelAutomator
from export_manager import execute_print_merge_to_pdf
from session_manager import cleanup_old_sessions
from config_manager import load_config

def parse_args():
    # Deprecated by JSON Bridge, retained structurally if needed.
    parser = argparse.ArgumentParser(description='PlakaMatik Python Export Hand-off Engine')
    return parser.parse_known_args()[0]

# Initialize Logging (Session ID generated natively)
init_logger(config.SESSION_ID, config.LOGS_DIR)

def run_pipeline(args):
    print(f"--- Starting LTO Automation Batch (Session {config.SESSION_ID}) ---")
    
    # 2. JSON Bridge Configuration Integration
    dynamic_config = load_config()
    is_cmyk = dynamic_config["COLOR_MODE"].upper() == "CMYK"
    is_visible = dynamic_config["CORELDRAW_VISIBLE"]
    global_dx = dynamic_config["GLOBAL_OFFSETS"].get("dx", 0.0)
    global_dy = dynamic_config["GLOBAL_OFFSETS"].get("dy", 0.0)

    # 3. Process the data enforcing max batch rules
    data_records = parse_input_data(config.INPUT_TXT_PATH)
    
    if not data_records:
        print("Pipeline stopped: Data processing failed.")
        return

    # 2. Cleanup temp files in Outputs and temp_previews not belonging to this active session
    cleanup_old_sessions([config.OUTPUTS_DIR, config.TEMP_PREVIEWS_DIR], config.SESSION_ID)
        
    final_pdf_path = os.path.join(config.OUTPUTS_DIR, f"LTO_Batch_{config.SESSION_ID}.pdf")

    # 3. Initialize the automation engine natively
    automator = CorelAutomator()
    automator.bypass_trial_screen()
    
    # 4. Connect to Corel DRAW securely via pipeline
    if automator.connect():
        # Override physical hardware visual state globally based on Flutter UI Setting
        try:
            automator.corel.Visible = is_visible
        except Exception as ve:
            print(f"Hardware bypass warning: Failed setting visibility property: {ve}")

        if os.path.exists(final_pdf_path):
            try:
                os.remove(final_pdf_path)
            except Exception as e:
                print(f"Warning: Could not delete old PDF. Close it if it's open. {e}")

        # =====================================================================================
        # ATOMIC FAILURE HANDLING & THE "DRAWBRIDGE" COM AUTOMATION INTERLOCK
        # =====================================================================================
        try:
            print("Initiating PDF Export Stage...")
            time.sleep(1)

            merge_success = execute_print_merge_to_pdf(
                automator.corel,
                data_records,
                final_pdf_path,
                config.TEMPLATE_MV_PATH,
                config.TEMPLATE_MC_PATH,
                force_cmyk=is_cmyk,
                global_dx=global_dx,
                global_dy=global_dy
            )

            if merge_success:
                print(f"--- Pipeline complete. Previews generated: {final_pdf_path} ---")
            else:
                print("--- Pipeline failed during template orchestration ---")
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Critical execution error: {e}")
        finally:
            try:
                for i in range(1, automator.corel.Documents.Count + 1):
                    automator.corel.Documents.Item(i).Dirty = False 
                    automator.corel.Documents.Item(i).Close()       
                print("All orphaned CorelDRAW COM handles released safely from memory.")
            except:
                pass

if __name__ == "__main__":
    args = parse_args()
    print(f"Starting pipeline with Config -> CMYK:{args.cmyk} VISIBLE:{args.visible}")
    run_pipeline(args)
