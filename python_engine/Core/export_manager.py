import os
import shutil
import time
import traceback
from text_mapper import replace_text_in_shapes

# A3 landscape dimensions (centimetres)
A3_W_CM = 42.0
A3_H_CM = 29.7


def _heartbeat(label=""):
    """Print a heartbeat so Flutter's process listener knows we're alive."""
    print(f"[HEARTBEAT] {label}" if label else "[HEARTBEAT] Engine running...", flush=True)


def _inject_text(corel_app, cdr_path, record_data, p_type):
    """
    Opens cdr_path, injects text into Print Layer IN-MEMORY.
    Returns (doc, page). Caller must close the doc.
    """
    print(f"[STAGE] Opening: {os.path.basename(cdr_path)} ({p_type})")
    doc = corel_app.OpenDocument(cdr_path)
    doc.Unit = 4  # cdrCentimeter
    page = doc.Pages.Item(1)

    mid = record_data.get('middle', '')
    iid = record_data.get('identifier', '')
    print(f"[STAGE] Injecting -> middle='{mid}' identifier='{iid}'")

    for i in range(1, page.Layers.Count + 1):
        lyr = page.Layers.Item(i)
        if lyr.Name.upper() == "PRINT LAYER":
            replace_text_in_shapes(lyr.Shapes, record_data, p_type)
            break

    _heartbeat("Text injection complete")
    return doc, page


def _export_pdfs_from_doc(doc, preview_path, print_path):
    """Export the active document to PREVIEW and PRINT PDFs."""
    doc.ClearSelection()
    page = doc.Pages.Item(1)
    pdf_settings = doc.PDFSettings
    pdf_settings.PublishRange = 0
    pdf_settings.ColorMode = 1   # RGB

    # ── PREVIEW export: Print Layer + Mock Layer visible, Guides hidden ──
    print("[STAGE] Exporting PREVIEW PDF...")
    for li in range(1, page.Layers.Count + 1):
        lyr = page.Layers.Item(li)
        lu  = lyr.Name.upper()
        try:
            if "GUIDE" in lu:
                lyr.Printable = False;  lyr.Visible = False
            else:
                lyr.Printable = True;   lyr.Visible = True
        except:
            pass
    doc.PublishToPDF(preview_path)
    print(f"Preview PDF saved: {preview_path}")
    _heartbeat("Preview exported")

    # ── PRINT export: Print Layer only, Mock Layer + Guides hidden ──
    print("[STAGE] Exporting PRINT PDF...")
    for li in range(1, page.Layers.Count + 1):
        lyr = page.Layers.Item(li)
        lu  = lyr.Name.upper()
        try:
            if "PRINT" in lu and "MOCK" not in lu:
                lyr.Printable = True;  lyr.Visible = True
            else:
                lyr.Printable = False; lyr.Visible = False
        except:
            pass
    doc.PublishToPDF(print_path)
    print(f"Print PDF saved: {print_path}")
    _heartbeat("Print exported")


def execute_print_merge_to_pdf(corel_app, data_records, output_pdf_path,
                               template_mv_path, template_mc_path,
                               global_dx=0.0, global_dy=0.0):
    """
    Entry point.
    Native CorelDRAW Master Canvas approach:
    - No pypdf overlay.
    - Resolves all text-reversion bugs by saving temp injected files to disk.
    - Safe from page-resize coordinate offset bugs.
    """
    try:
        total_records = len(data_records)
        print(f"[STAGE] Initializing Master Engine. Total records: {total_records}")
        _heartbeat("Engine started")

        if total_records == 0:
            return False

        base         = output_pdf_path.replace(".pdf", "")
        preview_path = f"{base}_PREVIEW.pdf"
        print_path   = f"{base}_PRINT.pdf"

        # ── 1. Create temporary injected files and SAVE to disk ─────────
        # This locks in the text values so they survive any copy-paste!
        temp_files = []
        for p_idx, record_data in enumerate(data_records):
            p_type = record_data.get("type", "MV").upper()
            template_path = template_mc_path if p_type == "MC" else template_mv_path

            if not os.path.exists(template_path):
                print(f"Error: Template not found for record {p_idx+1}")
                continue

            temp_cdr = template_path.replace('.cdr', f'__temp_locked_{p_idx}.cdr')
            shutil.copy2(template_path, temp_cdr)

            doc, page = _inject_text(corel_app, temp_cdr, record_data, p_type)
            # CRITICAL: Save to disk so clipboard reads the new values
            doc.Save()
            doc.Close()
            temp_files.append((temp_cdr, p_type))
            _heartbeat(f"Record {p_idx+1} locked to disk")

        if not temp_files:
            return False

        # ── 2. Create the unified A3 Master Document ─────────
        print("[STAGE] Constructing Native A3 Master Document...")
        master_doc = corel_app.CreateDocument()
        master_doc.Unit = 4
        master_page = master_doc.Pages.Item(1)
        master_page.SetSize(A3_W_CM, A3_H_CM)

        master_print_lyr = master_page.CreateLayer("Print Layer")
        master_mock_lyr  = master_page.CreateLayer("Mock Layer")
        _heartbeat("A3 Master ready")

        # ── 3. Paste and position each plate onto the Master ─────────
        for p_idx, (temp_cdr, p_type) in enumerate(temp_files):
            src_doc = corel_app.OpenDocument(temp_cdr)
            src_page = src_doc.Pages.Item(1)
            print(f"[STAGE] Merging plate {p_idx+1}/{len(temp_files)}...")

            for li in range(1, src_page.Layers.Count + 1):
                src_lyr = src_page.Layers.Item(li)
                lu = src_lyr.Name.upper()
                if "GUIDE" in lu or src_lyr.Shapes.Count == 0:
                    continue

                sr = corel_app.CreateShapeRange()
                for si in range(1, src_lyr.Shapes.Count + 1):
                    s = src_lyr.Shapes.Item(si)
                    if s.Type != 9:
                        sr.Add(s)

                if sr.Count == 0: continue

                # Copy shapes
                if sr.Count > 1:
                    grp = sr.Group()
                    grp.Copy()
                else:
                    sr.Item(1).Copy()

                time.sleep(0.3)

                # Paste into master
                master_doc.Activate()
                if "MOCK" in lu:
                    master_mock_lyr.Activate()
                else:
                    master_print_lyr.Activate()

                pasted = master_doc.ActiveLayer.Paste()
                time.sleep(0.3)

                # Center of the pasted group
                pcx = pasted.PositionX + (pasted.SizeWidth / 2.0)
                pcy = pasted.PositionY - (pasted.SizeHeight / 2.0)

                # Positioning Logic
                target_cx = (A3_W_CM / 2.0) + global_dx

                if len(temp_files) == 1:
                    # Single plate -> Center of A3
                    target_cy = (A3_H_CM / 2.0) + global_dy
                else:
                    # Batch -> Record 0 Top half, Record 1 Bottom half
                    target_cy = (A3_H_CM * 0.75 if p_idx == 0 else A3_H_CM * 0.25) + global_dy

                pasted.Move(target_cx - pcx, target_cy - pcy)

            src_doc.Dirty = False
            src_doc.Close()
            _heartbeat(f"Plate {p_idx+1} positioned")

        # ── 4. Export and Cleanup ─────────
        master_doc.Activate()
        _export_pdfs_from_doc(master_doc, preview_path, print_path)

        master_doc.Dirty = False
        master_doc.Close()

        for temp_cdr, _ in temp_files:
            try:
                os.remove(temp_cdr)
            except:
                pass

        print(f"Silenced PDF exports achieved. {len(temp_files)}/{total_records} records exported.")
        return True

    except Exception as e:
        print(f"Print Merge Error: {e}")
        traceback.print_exc()
        return False
