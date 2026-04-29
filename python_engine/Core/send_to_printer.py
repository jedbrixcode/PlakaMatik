import sys
import os
import time
import win32print
import win32api


# ── Helpers ───────────────────────────────────────────────────────────────────

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
                return True          # file opened → no lock
        except (PermissionError, OSError):
            print(f"[SPOOL] File locked — waiting ({attempt}/{retries})...")
            time.sleep(interval)
    print("[SPOOL] File still locked after max retries. Aborting.")
    return False


def check_printer_offline(printer_name: str) -> bool:
    """Return True if the printer is offline, errored, or unreachable."""
    try:
        phandle = win32print.OpenPrinter(printer_name)
        info    = win32print.GetPrinter(phandle, 2)
        win32print.ClosePrinter(phandle)
        status  = info['Status']
        # 0x80 = OFFLINE  |  0x4000 = NOT_RESPONDING  |  0x02 = ERROR
        return bool(status & 0x00000080 or status & 0x00004000 or status & 0x00000002)
    except Exception as e:
        print(f"[SPOOL] Hardware check error: {e}")
        return True   # treat unknown state as offline for safety


# ── Tier 1: Direct win32print RAW spooling ────────────────────────────────────

def _spool_via_win32print(pdf_path: str, printer_name: str) -> bool:
    """Send the PDF bytes directly to the Windows Print Spooler as a RAW job.

    This method bypasses ShellExecute entirely — no PDF viewer, no UAC dialog,
    no 'Access is denied' from shell handler associations.
    Returns True on success, False if the printer doesn't accept RAW streams
    (some GDI-only printers reject RAW; fallback to tier 2 in that case).
    """
    try:
        print(f"[SPOOL-T1] Opening printer handle: {printer_name}")
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, (
                f"PlakaMatik_{os.path.basename(pdf_path)}",  # job name
                None,                                         # output file
                "RAW"                                         # datatype
            ))
            try:
                win32print.StartPagePrinter(hPrinter)
                with open(pdf_path, 'rb') as f:
                    data = f.read()
                win32print.WritePrinter(hPrinter, data)
                win32print.EndPagePrinter(hPrinter)
                print(f"[SPOOL-T1] {len(data):,} bytes written to spooler.")
                return True
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        print(f"[SPOOL-T1] RAW spool failed: {e}  — trying fallback...")
        return False


# ── Tier 2: ShellExecute 'printto' fallback ───────────────────────────────────

def _spool_via_shell(pdf_path: str, printer_name: str) -> bool:
    """Use ShellExecute with the 'printto' verb as a fallback.

    'printto' sends directly to a named printer without opening a viewer window,
    and is more reliable than 'print' (which uses the default printer).
    """
    try:
        print(f"[SPOOL-T2] ShellExecute 'printto' → {printer_name}")
        win32api.ShellExecute(
            0,
            "printto",
            pdf_path,
            f'"{printer_name}"',
            ".",
            0
        )
        # Give the shell handler 3 s to absorb the spool request
        time.sleep(3)
        print("[SPOOL-T2] ShellExecute submitted successfully.")
        return True
    except Exception as e:
        print(f"[SPOOL-T2] ShellExecute fallback also failed: {e}")
        return False


# ── Public API ────────────────────────────────────────────────────────────────

def print_pdf(pdf_path: str, printer_name: str, job_type: str = "single"):
    """Main entry point called by main.py.

    Strategy:
      1. Normalize the path (handles 'Win10 PRO' spaces in profile name).
      2. Wait until the file is no longer write-locked by CorelDRAW.
      3. Verify the printer is online.
      4. Try Tier-1 (direct win32print RAW).
      5. Fall back to Tier-2 (ShellExecute 'printto').
      6. Hard-fail with a clear ACCESS DENIED message if both tiers fail.
    """
    # 1. Normalize
    pdf_path = _normalize(pdf_path)
    log_type = "batch dual-plates" if job_type == "multiple" else "single print"
    print(f"[SPOOL] Target   : {pdf_path}")
    print(f"[SPOOL] Printer  : {printer_name}")
    print(f"[SPOOL] Job type : {log_type}")

    # 2. File-lock guard — wait for CorelDRAW to release the PDF handle
    if not _wait_for_file_release(pdf_path):
        err = "ACCESS IS DENIED: PDF is still locked by another process."
        print(err, file=sys.stderr)
        print(err)
        sys.exit(1)

    # 3. Printer online check
    print(f"[SPOOL] Checking printer status...")
    if check_printer_offline(printer_name):
        err = (
            "HALTING: PRINTER IS OFFLINE, ERRORED, OR NOT RESPONDING.\n"
            "Please check if the Canon iX6700 is online and set as 'Shared'."
        )
        print(err)
        print(err, file=sys.stderr)
        sys.exit(1)

    # 4. Tier 1 — direct win32print RAW
    if _spool_via_win32print(pdf_path, printer_name):
        print(f"[SPOOL] Success — {log_type} spooled via win32print.")
        return

    # 5. Tier 2 — ShellExecute 'printto'
    if _spool_via_shell(pdf_path, printer_name):
        print(f"[SPOOL] Success — {log_type} spooled via ShellExecute.")
        return

    # 6. Both tiers failed
    err = (
        "ACCESS IS DENIED: Both spooling methods failed.\n"
        "Check that the Canon iX6700 is online and set as a 'Shared' printer,\n"
        "or run PlakaMatik as Administrator."
    )
    print(err)
    print(err, file=sys.stderr)
    sys.exit(1)


# ── CLI entry point ───────────────────────────────────────────────────────────

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

    print_pdf(target_pdf, target_printer, job_type)

