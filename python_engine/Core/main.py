import os
import sys
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
from send_to_printer import print_pdf

def parse_args():
    parser = argparse.ArgumentParser(description='PlakaMatik Python Export Hand-off Engine')
    parser.add_argument('--action', type=str, default='generate', choices=['generate', 'spool'], help='Action to perform')
    parser.add_argument('--pdf', type=str, default=None, help='Target PDF path for spooling')
    parser.add_argument('--config', type=str, default=None, help='Absolute path to config.json')
    parser.add_argument('--session-id', type=str, default=None, help='Override the session ID timestamp')
    return parser.parse_known_args()[0]

# Initialize Logging (Session ID generated natively)
init_logger(config.SESSION_ID, config.LOGS_DIR)

def run_pipeline(args):
    print(f"--- Starting LTO Automation Batch (Session {config.SESSION_ID}) ---")
    
    # 2. JSON Bridge Configuration Integration
    dynamic_config = load_config(args.config)
    is_visible = dynamic_config["CORELDRAW_VISIBLE"]
    global_dx = dynamic_config["GLOBAL_OFFSETS"].get("dx", 0.0)
    global_dy = dynamic_config["GLOBAL_OFFSETS"].get("dy", 0.0)
    printer_name = dynamic_config.get("PRINTER_NAME", "Microsoft Print to PDF")

    # 3. Process the data enforcing max batch rules
    data_records = parse_input_data(config.INPUT_TXT_PATH)
    
    if not data_records:
        print("Pipeline stopped: Data processing failed.")
        return

    # 2. Cleanup temp files in Outputs and temp_previews not belonging to this active session
    cleanup_old_sessions([config.OUTPUTS_DIR, config.TEMP_PREVIEWS_DIR], config.SESSION_ID)
        
    final_pdf_path = os.path.join(config.OUTPUTS_DIR, f"{config.SESSION_ID}.pdf")

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
                global_dx=global_dx,
                global_dy=global_dy
            )

            if merge_success:
                print(f"--- Pipeline complete. Previews generated: {final_pdf_path} ---")
                print("Waiting for operator to confirm physical print action...")
            else:
                err_msg = "Pipeline failed during template orchestration."
                print(f"--- {err_msg} ---")
                print(err_msg, file=sys.stderr)
        except Exception as e:
            import traceback
            traceback.print_exc()
            err_msg = f"Critical execution error: {e}"
            print(err_msg)
            print(err_msg, file=sys.stderr)
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
    
    # Inject override for session ID if requested by Flutter
    if args.session_id:
        config.SESSION_ID = args.session_id
        
    # Re-init logger since session ID might have changed
    init_logger(config.SESSION_ID, config.LOGS_DIR)
        
    if args.action == 'spool':
        print(f"Orchestrator received manual spool request for: {args.pdf}")
        if not args.pdf or not os.path.exists(args.pdf):
            err_msg = f"HALTING: Cannot spool. Invalid or missing PDF path: {args.pdf}"
            print(err_msg, file=sys.stderr)
            sys.exit(1)
            
        dynamic_config = load_config(args.config)
        printer_name = dynamic_config.get("PRINTER_NAME", "Microsoft Print to PDF")
        
        print_pdf(args.pdf, printer_name, "manual")
    else:
        print(f"Starting Unified Orchestrator Pipeline...")
        run_pipeline(args)
