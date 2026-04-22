import os
import time

# Custom Modules
import config
from engine_logger import init_logger
from data_processor import parse_input_data
from corel_engine import CorelAutomator
from print_handler import execute_print_merge_to_pdf
from session_manager import cleanup_old_sessions

# Initialize Logging
init_logger(config.SESSION_ID, config.LOGS_DIR)

def run_pipeline():
    print(f"--- Starting LTO Automation Batch (Session {config.SESSION_ID}) ---")
    
    # 1. Process the data enforcing max batch rules
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
    
    # 4. Connect to Corel DRAW securely via headless-compatible approach
    if automator.connect():
        # Prevent locked file crashes by attempting to gently override
        if os.path.exists(final_pdf_path):
            try:
                os.remove(final_pdf_path)
            except Exception as e:
                print(f"Warning: Could not delete old PDF. Close it if it's open. {e}")

        # =====================================================================================
        # ATOMIC FAILURE HANDLING & THE "DRAWBRIDGE" COM AUTOMATION INTERLOCK
        # =====================================================================================
        # Properly releasing the COM object in a finally block ensures that background CorelDRAW 
        # instances are terminated even if a crash occurs, preventing "ghost" processes from 
        # consuming all system RAM and eventually crashing the entire machine.
        #
        # Wrapping entire orchestration logic inside a "try / except / finally" block,
        # 1. TRY: Run the dangerous logic.
        # 2. EXCEPT: If it fails, catch the error.
        # 3. FINALLY: No matter what happened (SUCCESS or FAILURE), forcefully run the 
        #    cleanup scripts to kill hanging document handles.
        # =====================================================================================
        try:
            print("Initiating printing logic sequence...")
            time.sleep(2)

            merge_success = execute_print_merge_to_pdf(
                automator.corel,
                data_records,
                final_pdf_path,
                config.TEMPLATE_MV_PATH,
                config.TEMPLATE_MC_PATH
            )

            if merge_success:
                print(f"--- Pipeline complete. Preview generated: {final_pdf_path} ---")
            else:
                print("--- Pipeline failed during template orchestration ---")
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Critical execution error: {e}")
        finally:
            # ---------------------------------------------------------------------------------
            # THE HARDWARE/MEMORY RESET
            # ---------------------------------------------------------------------------------
            # The finally block serves as a fail-safe mechanism that prevents memory leaks by 
            # forcefully closing all active CorelDRAW documents without saving, ensuring that 
            # crashed or "ghost" processes are cleared and templates remain unaltered.
            # ---------------------------------------------------------------------------------
            try:
                for i in range(1, automator.corel.Documents.Count + 1):
                    automator.corel.Documents.Item(i).Dirty = False # Prevent Save prompt
                    automator.corel.Documents.Item(i).Close()       # Kill document handle
                print("All orphaned CorelDRAW COM handles released safely from memory.")
            except:
                pass

if __name__ == "__main__":
    run_pipeline()
