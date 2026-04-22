import sys
import os

class ConsoleLogger:
    """
    Reroutes pipeline outputs directly to physical text files
    for the Flutter StreamUI rather than the invisible backend shell.
    """
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

def init_logger(session_id, logs_dir):
    """
    Injects the hook to overwrite the native sys.stdout.
    """
    log_file = os.path.join(logs_dir, f"Log_{session_id}.txt")
    sys.stdout = ConsoleLogger(log_file, session_id)
