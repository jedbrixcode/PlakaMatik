import sys
import os

class ConsoleLogger:
    """
    Reroutes pipeline outputs directly to physical text files
    for the Flutter StreamUI rather than the invisible backend shell.
    """
    def __init__(self, terminal, log_path, session_id, prefix=""):
        self.terminal = terminal
        self.log_path = log_path
        self.session_id = session_id
        self.prefix = prefix
        
        # Write a session start marker only on stdout initialization
        if not prefix:
            try:
                with open(self.log_path, "a", encoding="utf-8") as f:
                    f.write(f"\n\n--- [SESSION START: {session_id}] ---\n")
            except:
                pass

    def write(self, message):
        if self.terminal is not None:
            try:
                self.terminal.write(message)
            except:
                pass
                
        try:
            if message.strip() or message == '\n':
                with open(self.log_path, "a", encoding="utf-8") as f:
                    if self.prefix and message.strip():
                        f.write(f"{self.prefix}{message}")
                    else:
                        f.write(message)
        except:
            pass

    def flush(self):
        if self.terminal is not None:
            try:
                self.terminal.flush()
            except:
                pass

def init_logger(session_id, logs_dir):
    """
    Injects the hook to overwrite the native sys.stdout and sys.stderr.
    """
    log_file = os.path.join(logs_dir, f"Log_{session_id}.txt")
    sys.stdout = ConsoleLogger(sys.stdout, log_file, session_id)
    sys.stderr = ConsoleLogger(sys.stderr, log_file, session_id, prefix="[ERROR] ")
