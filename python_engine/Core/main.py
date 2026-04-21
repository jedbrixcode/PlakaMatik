import os
import time
import sys
import glob
from datetime import datetime
from data_processor import parse_input_data
from corel_engine import CorelAutomator
from print_handler import execute_print_merge_to_pdf

# System paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
FLUTTER_ROOT_DIR = os.path.dirname(ROOT_DIR)

# Subfolder for Templates
PLATE_TEMPLATE_DIR = os.path.join(ROOT_DIR, "CorelDRAW Templates", "Main Templates")

# Dynamic routing to the correct folders
PLAKAMATIK_DIR = os.path.join(os.path.expanduser("~"), "Documents", "PlakaMatik Files")
INPUT_TXT_PATH = os.path.join(PLAKAMATIK_DIR, "csv", "flutter_user_input.txt")

TEMPLATE_MV_PATH = os.path.join(PLATE_TEMPLATE_DIR, "MV_PLATE.cdr")
TEMPLATE_MC_PATH = os.path.join(PLATE_TEMPLATE_DIR, "MC_PLATE.cdr")

# Generate Session ID (yyyyMMdd_HHmm)
SESSION_ID = datetime.now().strftime("%Y%m%d_%H%M")
LOGS_DIR = os.path.join(PLAKAMATIK_DIR, "Logs")
if not os.path.exists(LOGS_DIR):
    os.makedirs(LOGS_DIR)

class ConsoleLogger:
    def __init__(self, log_path, session_id):
        self.terminal = sys.stdout
        self.log_path = log_path
        self.session_id = session_id
        
        # Write a session start marker
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(f"\n\n--- [SESSION START: {session_id}] ---\n")

    def write(self, message):
        self.terminal.write(message)
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(message)
        except:
            pass

    def flush(self):
        self.terminal.flush()

# Redirect stdout to pipe terminal logs to a session specific file
sys.stdout = ConsoleLogger(os.path.join(LOGS_DIR, f"Log_{SESSION_ID}.txt"), SESSION_ID)

def cleanup_old_sessions(directories, current_session):
    """
    Purges ghost artifact files from previous instances locking them to maintain UI performance.
    """
    print(f"Running cleanup. Protecting current session: {current_session}")
    for directory in directories:
        if not os.path.exists(directory):
            continue
        for file in glob.glob(os.path.join(directory, "*")):
            filename = os.path.basename(file)
            if current_session not in filename:
                try:
                    # Ignore .gitignores if present
                    if not filename.startswith("."): 
                        os.remove(file)
                        print(f"Purged old ghost file: {filename}")
                except Exception as e:
                    print(f"Warning: Could not delete {filename}. {e}")

def run_pipeline():
    print(f"--- Starting LTO Automation Batch (Session {SESSION_ID}) ---")
    
    # 1. Process the data enforcing max batch rules
    data_records = parse_input_data(INPUT_TXT_PATH)
    
    if not data_records:
        print("Pipeline stopped: Data processing failed.")
        return

    # 2. Cleanup temp files in Outputs and temp_previews not belonging to this active session
    outputs_dir = os.path.join(PLAKAMATIK_DIR, "Outputs")
    temp_previews_dir = os.path.join(PLAKAMATIK_DIR, "temp_previews")
    if not os.path.exists(outputs_dir): os.makedirs(outputs_dir)
    if not os.path.exists(temp_previews_dir): os.makedirs(temp_previews_dir)
    
    cleanup_old_sessions([outputs_dir, temp_previews_dir], SESSION_ID)
        
    final_pdf_path = os.path.join(outputs_dir, f"LTO_Batch_{SESSION_ID}.pdf")

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
        # Why is this so crucial? 
        # When we connect to CorelDRAW via win32com (COM API), we are opening invisible 
        # background instances of the CorelDRAW application in your Windows Memory (RAM).
        # If the Python script crashes during execution (e.g. malformed text, missing template), 
        # and we DO NOT explicitly close the document and release the COM object application lock,
        # an invisible "ghost" copy of CorelDRAW will remain running in the background forever.
        #
        # If this happens multiple times over a workday, the computer will eventually run out 
        # of RAM completely, causing the printer and system to crash drastically.
        # 
        # By wrapping our entire orchestration logic inside a "try / except / finally" block,
        # we create an "Atomic" constraint:
        # 1. TRY: Run the dangerous logic (opening, pasting, exporting).
        # 2. EXCEPT: If it fails, catch the error (don't let Python just die!).
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
                TEMPLATE_MV_PATH,
                TEMPLATE_MC_PATH
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
            # This 'finally' block ensures that even if a "ghost" process crashes one run, 
            # the memory state is reset. We iterate backwards or forwards through the 
            # absolute live document count of the CorelDRAW instance and force close them 
            # without saving (.Dirty = False) so templates aren't destructively overwritten.
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
