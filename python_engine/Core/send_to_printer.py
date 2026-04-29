import sys
import os
import time

def _normalize(path: str) -> str:
    """Return an absolute, backslash-normalized path — safe for Win32 APIs
    even when the user-profile name contains a space (e.g. 'Win10 PRO')."""
    return os.path.normpath(os.path.abspath(path))

def _wait_for_file_release(path: str, retries: int = 10, interval: float = 1.0) -> bool:
    """Retry-loop that waits until the file is no longer write-locked.
    CorelDRAW may hold an exclusive handle on the PDF for a short time
    after export. Returns True when the file is accessible, False on timeout.
    """
    for attempt in range(1, retries + 1):
        try:
            with open(path, 'rb'):
                return True
        except (PermissionError, OSError):
            print(f"[PRINT] File locked — waiting ({attempt}/{retries})...")
            time.sleep(interval)
    print("[PRINT] File still locked after max retries. Aborting.")
    return False

def _select_printer(corel_doc, printer_name: str) -> bool:
    """
    CorelDRAW exposes printer destination selection via PrintSettings.SelectPrinter(name).
    Setting PrintSettings.Printer directly can be read-only depending on the CorelDRAW COM type library.
    """
    ps = corel_doc.PrintSettings
    # Preferred: method-based selection (SDK: PrintSettings.SelectPrinter)
    try:
        if hasattr(ps, "SelectPrinter"):
            ps.SelectPrinter(printer_name)
            print(f"[PRINT] Printer selected via SelectPrinter(): {printer_name}")
            return True
    except Exception as e:
        print(f"[PRINT] SelectPrinter failed: {e}")

    # Fallback: property assignment (may still be supported on some installs)
    try:
        ps.Printer = printer_name
        print(f"[PRINT] Printer selected via PrintSettings.Printer: {printer_name}")
        return True
    except Exception as e:
        print(f"[PRINT] PrintSettings.Printer set failed: {e}")

    # Last resort: set Windows default printer (so CorelDRAW prints to the default destination)
    try:
        import win32print
        win32print.SetDefaultPrinter(printer_name)
        print(f"[PRINT] Set Windows default printer to: {printer_name}")
        return True
    except Exception as e:
        print(f"[PRINT] Could not set Windows default printer: {e}")

    return False

def print_pdf_via_corel(pdf_path: str, printer_name: str, job_type: str = "single"):
    """
    Prints the composed A3 PDF master file directly through CorelDRAW.
    This replaces the unreliable Win32 native spooling.
    """
    pdf_path = _normalize(pdf_path)
    log_type = "batch dual-plates" if job_type == "multiple" else "single print"
    print(f"[PRINT] Target   : {pdf_path}")
    print(f"[PRINT] Printer  : {printer_name}")
    print(f"[PRINT] Job type : {log_type}")

    if not _wait_for_file_release(pdf_path):
        err = "ACCESS IS DENIED: PDF is still locked by another process."
        print(err, file=sys.stderr)
        print(err)
        sys.exit(1)

    from corel_engine import CorelAutomator
    automator = CorelAutomator()
    # Bypass delay set to 0 because CorelDRAW is likely already open and bypassed 
    # during the generate phase, or we don't want to wait 5s for every print.
    automator.bypass_trial_screen(delay=0)
    
    if not automator.connect():
        err = "HALTING: Could not connect to CorelDRAW to print."
        print(err, file=sys.stderr)
        print(err)
        sys.exit(1)

    try:
        print(f"[PRINT] Opening A3 Master file in CorelDRAW...")
        doc = automator.corel.OpenDocument(pdf_path)
        
        # Configure CorelDRAW print settings
        print(f"[PRINT] Configuring printer: {printer_name}")
        if not _select_printer(doc, printer_name):
            print(
                "[PRINT] WARNING: Could not select the requested printer in CorelDRAW. "
                "Proceeding with CorelDRAW's current/default printer."
            )
        
        # Explicit CMYK calibration
        doc.PrintSettings.ColorMode = 2  # prnColorCMYK = 2
        
        # Explicit A3 paper size (prnPaperA3 = 8)
        doc.PrintSettings.PaperSize = 8

        print(f"[PRINT] Submitting to hardware spooler via CorelDRAW...")
        doc.PrintOut()
        
        # Give CorelDRAW a brief moment to finish spooling the job to the OS
        time.sleep(3)
        
        doc.Close()
        print(f"[PRINT] Success — {log_type} printed natively via CorelDRAW.")
        
    except Exception as e:
        err = f"CorelDRAW Printing Error: {e}"
        print(err, file=sys.stderr)
        print(err)
        try:
            doc.Close()
        except:
            pass
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python send_to_printer.py <pdf_path> <printer_name> [job_type]")
        sys.exit(1)

    target_pdf     = sys.argv[1]
    target_printer = sys.argv[2]
    job_type       = sys.argv[3] if len(sys.argv) > 3 else "single"

    try:
        from config_manager import load_config
        dyn_conf = load_config()
        if dyn_conf.get("PRINTER_NAME"):
            target_printer = dyn_conf["PRINTER_NAME"]
    except Exception:
        pass

    if not os.path.exists(target_pdf):
        print(f"HALTING: Payload PDF does not exist: {target_pdf}")
        sys.exit(1)

    print_pdf_via_corel(target_pdf, target_printer, job_type)
