import os
import time
import traceback
from text_mapper import replace_text_in_shapes


def _heartbeat(label=""):
    """Print a heartbeat so Flutter's process listener knows the engine is alive."""
    print(f"[HEARTBEAT] {label}" if label else "[HEARTBEAT] Engine running...", flush=True)


def _set_layer_visibility(page, layer_name_upper, printable, visible):
    """Helper: set printable and visible for a named layer on a CorelDRAW page."""
    for i in range(1, page.Layers.Count + 1):
        lyr = page.Layers.Item(i)
        if lyr.Name.upper() == layer_name_upper:
            try:
                lyr.Printable = printable
                lyr.Visible = visible
            except Exception as e:
                print(f"Warning: Could not set layer '{lyr.Name}' visibility: {e}")
            return
    print(f"Note: Layer '{layer_name_upper}' not present on this page.")


def _inject_text(corel_app, template_path, record_data, p_type):
    """
    Opens a CDR template, injects text into the Print Layer.
    Returns (doc, page). Caller is responsible for closing the doc.
    """
    print(f"[STAGE] Opening template: {os.path.basename(template_path)} ({p_type})")
    doc = corel_app.OpenDocument(template_path)
    doc.Unit = 4  # cdrCentimeter
    page = doc.Pages.Item(1)

    middle_val     = record_data.get('middle', '')
    identifier_val = record_data.get('identifier', '')
    print(f"[STAGE] Injecting -> middle='{middle_val}' identifier='{identifier_val}'")

    for i in range(1, page.Layers.Count + 1):
        lyr = page.Layers.Item(i)
        if lyr.Name.upper() == "PRINT LAYER":
            replace_text_in_shapes(lyr.Shapes, record_data, p_type)
            break

    _heartbeat("Text injection complete")
    return doc, page


def _export_single_record(corel_app, record_data, template_path,
                          preview_pdf_path, print_pdf_path,
                          global_dx=0.0, global_dy=0.0):
    """
    Single-plate path: open template, inject text, toggle layers, export two PDFs.
    """
    p_type = record_data.get("type", "MV").upper()

    if not os.path.exists(template_path):
        print(f"Error: Template not found -> {template_path}")
        return False

    doc, page = _inject_text(corel_app, template_path, record_data, p_type)

    doc.ClearSelection()
    pdf_settings = doc.PDFSettings
    pdf_settings.PublishRange = 0
    pdf_settings.ColorMode = 1  # RGB

    print("[STAGE] Exporting PREVIEW PDF...")
    _set_layer_visibility(page, "PRINT LAYER", printable=True,  visible=True)
    _set_layer_visibility(page, "MOCK LAYER",  printable=True,  visible=True)
    _set_layer_visibility(page, "GUIDES",      printable=False, visible=False)
    doc.PublishToPDF(preview_pdf_path)
    print(f"Preview PDF saved: {preview_pdf_path}")
    _heartbeat("Preview exported")

    print("[STAGE] Exporting PRINT PDF...")
    _set_layer_visibility(page, "PRINT LAYER", printable=True,  visible=True)
    _set_layer_visibility(page, "MOCK LAYER",  printable=False, visible=False)
    _set_layer_visibility(page, "GUIDES",      printable=False, visible=False)
    doc.PublishToPDF(print_pdf_path)
    print(f"Print PDF saved: {print_pdf_path}")
    _heartbeat("Print PDF exported")

    doc.Dirty = False
    doc.Close()
    return True


def execute_print_merge_to_pdf(corel_app, data_records, output_pdf_path,
                               template_mv_path, template_mc_path,
                               global_dx=0.0, global_dy=0.0):
    """
    Entry point called by main.py.
      - 1 record  -> direct single-plate export
      - 2+ records -> build a fresh A3 master canvas; paste all plates positioned
                      top-half / bottom-half; export one PREVIEW and one PRINT PDF.
    """
    try:
        total_records = len(data_records)
        print(f"[STAGE] Initializing Master Engine. Total records: {total_records}")

        if total_records == 0:
            print("No records found.")
            return False

        base         = output_pdf_path.replace(".pdf", "")
        preview_path = f"{base}_PREVIEW.pdf"
        print_path   = f"{base}_PRINT.pdf"

        # ── SINGLE PLATE ────────────────────────────────────────────────────────
        if total_records == 1:
            record_data   = data_records[0]
            p_type        = record_data.get("type", "MV").upper()
            template_path = template_mc_path if p_type == "MC" else template_mv_path
            result = _export_single_record(
                corel_app, record_data, template_path,
                preview_path, print_path, global_dx, global_dy
            )
            ok = '1' if result else '0'
            print(f"Silenced PDF exports achieved. {ok}/1 records exported.")
            return result

        # ── BATCH (2+ records) -> fresh A3 master canvas ─────────────────────────
        print(f"[STAGE] Batch mode: building A3 master canvas for {total_records} plates...")
        _heartbeat("Starting batch canvas build")

        # A3 landscape: 420 mm × 297 mm = 42 cm × 29.7 cm
        A3_W_CM = 42.0
        A3_H_CM = 29.7

        print("[STAGE] Creating blank A3 master document...")
        master_doc  = corel_app.CreateDocument()
        master_doc.Unit = 4
        master_page = master_doc.Pages.Item(1)
        master_page.SetSize(A3_W_CM, A3_H_CM)

        # Build named layers on the master canvas
        try:
            master_print_lyr = master_page.Layers.Item(1)
            master_print_lyr.Name = "Print Layer"
        except:
            master_print_lyr = master_page.CreateLayer("Print Layer")

        try:
            master_mock_lyr = master_page.CreateLayer("Mock Layer")
        except:
            master_mock_lyr = master_page.Layers.Item("Mock Layer")

        _heartbeat("Master A3 canvas ready")

        # ── Paste each record's shapes onto the A3 canvas ──────────────────────
        for p_idx, record_data in enumerate(data_records):
            p_type        = record_data.get("type", "MV").upper()
            template_path = template_mc_path if p_type == "MC" else template_mv_path

            if not os.path.exists(template_path):
                print(f"Error: Template not found for record {p_idx+1}: {template_path}")
                continue

            src_doc, src_page = _inject_text(
                corel_app, template_path, record_data, p_type
            )
            _heartbeat(f"Merging record {p_idx+1}/{total_records}")

            # Vertical target: record 0 -> top half, record 1 -> bottom half
            target_cx = (A3_W_CM / 2.0) + global_dx
            target_cy = (A3_H_CM * 0.75 if p_idx == 0 else A3_H_CM * 0.25) + global_dy

            for li in range(1, src_page.Layers.Count + 1):
                src_lyr   = src_page.Layers.Item(li)
                lyr_upper = src_lyr.Name.upper()

                if "GUIDE" in lyr_upper or src_lyr.Shapes.Count == 0:
                    continue

                # Build a ShapeRange excluding guideline objects (Type 9)
                sr = corel_app.CreateShapeRange()
                for si in range(1, src_lyr.Shapes.Count + 1):
                    try:
                        s = src_lyr.Shapes.Item(si)
                        if s.Type != 9:
                            sr.Add(s)
                    except:
                        pass

                if sr.Count == 0:
                    continue

                # Copy the shapes (group first if multiple)
                try:
                    if sr.Count > 1:
                        grp = sr.Group()
                        grp.Copy()
                    else:
                        sr.Item(1).Copy()
                except Exception as copy_e:
                    print(f"  Warning: copy error on layer '{src_lyr.Name}': {copy_e}")
                    continue

                time.sleep(0.4)

                # Switch to master doc and activate the right destination layer
                master_doc.Activate()
                if "MOCK" in lyr_upper:
                    master_mock_lyr.Activate()
                else:
                    master_print_lyr.Activate()

                # Paste and reposition
                try:
                    pasted = master_doc.ActiveLayer.Paste()
                    time.sleep(0.4)
                    cx = pasted.PositionX + pasted.SizeWidth  / 2.0
                    cy = pasted.PositionY - pasted.SizeHeight / 2.0
                    pasted.Move(target_cx - cx, target_cy - cy)
                    print(f"  [MERGE {p_idx+1}] '{src_lyr.Name}' pasted & positioned.")
                except Exception as pos_e:
                    print(f"  Warning: paste/position error: {pos_e}")

                _heartbeat(f"Layer '{src_lyr.Name}' merged for record {p_idx+1}")

            src_doc.Dirty = False
            src_doc.Close()
            print(f"[STAGE] Record {p_idx+1} fully merged onto A3 canvas.")
            _heartbeat(f"Record {p_idx+1} complete")

        # ── Export PREVIEW (all layers visible, no Guides) ──────────────────────
        print("[STAGE] Exporting COMBINED VERIFICATION PREVIEW...")
        _heartbeat("Starting PREVIEW export")
        master_doc.ClearSelection()
        pdf_settings = master_doc.PDFSettings
        pdf_settings.PublishRange = 0
        pdf_settings.ColorMode    = 1

        for li in range(1, master_page.Layers.Count + 1):
            lyr = master_page.Layers.Item(li)
            lu  = lyr.Name.upper()
            if "GUIDE" in lu:
                lyr.Printable = False
                lyr.Visible   = False
            else:
                lyr.Printable = True
                lyr.Visible   = True

        master_doc.PublishToPDF(preview_path)
        print(f"Combined preview PDF saved: {preview_path}")
        _heartbeat("PREVIEW exported")

        # ── Export PRINT (Print Layer only, Mock hidden) ─────────────────────────
        print("[STAGE] Exporting COMBINED PRINT-READY PAYLOAD...")
        for li in range(1, master_page.Layers.Count + 1):
            lyr = master_page.Layers.Item(li)
            lu  = lyr.Name.upper()
            if "PRINT" in lu and "MOCK" not in lu:
                lyr.Printable = True
                lyr.Visible   = True
            else:
                lyr.Printable = False
                lyr.Visible   = False

        master_doc.PublishToPDF(print_path)
        print(f"Combined print PDF saved: {print_path}")
        _heartbeat("PRINT PDF exported")

        master_doc.Dirty = False
        master_doc.Close()
        print(f"Silenced PDF exports achieved. {total_records}/{total_records} records exported.")
        return True

    except Exception as e:
        print(f"Print Merge Error: {e}")
        traceback.print_exc()
        return False
