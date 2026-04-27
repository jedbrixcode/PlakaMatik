import sys
import os
import win32print
import win32api
import time

def check_printer_offline(printer_name):
    try:
        phandle = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(phandle, 2)
        win32print.ClosePrinter(phandle)
        status = info['Status']
        # 128 = PRINTER_STATUS_OFFLINE, 16384 = PRINTER_STATUS_NOT_RESPONDING, 2 = PRINTER_STATUS_ERROR
        if status & 0x00000080 or status & 0x00004000 or status & 0x00000002:
            return True
        return False
    except Exception as e:
        print(f"Hardware Error Checking Printer: {e}")
        return True

def print_pdf(pdf_path, printer_name, job_type="single"):
    print(f"Connecting to hardware spooler: {printer_name}")
    
    if check_printer_offline(printer_name):
        err_msg = "HALTING: PRINTER IS OFFLINE, ERRORED, OR NOT RESPONDING."
        print(err_msg)
        print(err_msg, file=sys.stderr)
        sys.exit(1)
        
    log_type = "batch dual-plates" if job_type == "multiple" else "single print"
    print(f"Transmitting {log_type} job...")
    
    try:
        win32api.ShellExecute(
            0,
            "print",
            pdf_path,
            f'/d:"{printer_name}"',
            ".",
            0
        )
        # Give OS 2 seconds to absorb the spool handle seamlessly
        time.sleep(2)
        print(f"Success! {log_type} spooled to device.")
    except Exception as e:
        err_msg = f"HALTING: Critical spool injection error: {e}"
        print(err_msg)
        print(err_msg, file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python send_to_printer.py <pdf_path> <printer_name> [job_type(single/multiple)]")
        sys.exit(1)
        
    target_pdf = sys.argv[1]
    target_printer = sys.argv[2]
    job_type = sys.argv[3] if len(sys.argv) > 3 else "single"
    
    try:
        from config_manager import load_config
        dyn_conf = load_config()
        if dyn_conf.get("PRINTER_NAME"):
            target_printer = dyn_conf["PRINTER_NAME"]
    except:
        pass
    
    if not os.path.exists(target_pdf):
        print(f"HALTING: Payload PDF {target_pdf} does not exist.")
        sys.exit(1)
        
    print_pdf(target_pdf, target_printer, job_type)
