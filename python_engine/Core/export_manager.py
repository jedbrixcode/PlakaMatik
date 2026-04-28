import os
import shutil
import time
import traceback
from text_mapper import replace_text_in_shapes


def _heartbeat(label=""):
    """Print a heartbeat so Flutter's process listener knows the engine is alive."""
    print(f"[HEARTBEAT] {label}" if label else "[HEARTBEAT] Engine running...", flush=True)


def _inject_text(corel_app, cdr_path, record_data, p_type):
    """
    Open cdr_path, inject text into Print Layer IN-MEMORY.
    Returns (doc, page). Caller must close.
    """
    print(f"[STAGE] Opening: {os.path.basename(cdr_path)} ({p_type})")
    doc  = corel_app.OpenDocument(cdr_path)
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


def _export_record_pdfs(corel_app, cdr_path, record_data, p_type,
                        preview_path, print_path):
    """
    THE ONLY RELIABLE APPROACH:
    Open cdr_path (a copy of the template), inject text IN-MEMORY,
    and export PREVIEW + PRINT PDFs directly from the SAME document.
    No cross-document copy-paste — injected values are guaranteed in output.
    """
    doc, page = _inject_text(corel_app, cdr_path, record_data, p_type)

    doc.ClearSelection()
    pdf_settings             = doc.PDFSettings
    pdf_settings.PublishRange = 0
    pdf_settings.ColorMode    = 1   # RGB

    # PREVIEW: Print Layer + Mock Layer visible, Guides hidden
    print("[STEP] Exporting PREVIEW...")
    for li in range(1, page.Layers.Count + 1):
        lyr = page.Layers.Item(li)
        lu  = lyr.Name.upper()
        try:
            if "GUIDE" in lu:
                lyr.Printable = False; lyr.Visible = False
            else:
                lyr.Printable = True;  lyr.Visible = True
        except:
            pass
    doc.PublishToPDF(preview_path)
    print(f"  -> {os.path.basename(preview_path)}")
    _heartbeat("Temp preview done")

    # PRINT: Print Layer only, Mock + Guides hidden
    print("[STEP] Exporting PRINT...")
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
    print(f"  -> {os.path.basename(print_path)}")
    _heartbeat("Temp print done")

    doc.Dirty = False
    doc.Close()
    return True


def _compose_a3(input_paths, output_path):
    """
    Compose multiple plate-sized PDFs onto a single A3 landscape page using
    pypdf merge_transformed_page (NOT merge_page which just overlays at origin).

    Positioning:
      1 plate  → horizontally and vertically centred on A3
      2 plates → plate[0] centred in top half, plate[1] centred in bottom half
    """
    try:
        from pypdf import PdfReader, PdfWriter, Transformation

        MM_TO_PT = 72.0 / 25.4
        a3_w = 420 * MM_TO_PT   # ~1190.55 pt  (landscape width)
        a3_h = 297 * MM_TO_PT   # ~841.89  pt  (landscape height)

        writer   = PdfWriter()
        a3_page  = writer.add_blank_page(width=a3_w, height=a3_h)
        n        = len(input_paths)

        for i, pdf_path in enumerate(input_paths):
            reader = PdfReader(pdf_path)
            plate  = reader.pages[0]
            pw     = float(plate.mediabox.width)
            ph     = float(plate.mediabox.height)

            # Horizontal centre on A3
            tx = (a3_w - pw) / 2.0

            if n == 1:
                # Single plate → vertical centre
                ty = (a3_h - ph) / 2.0
            elif i == 0:
                # First plate  → centre of TOP half
                ty = a3_h / 2.0 + (a3_h / 2.0 - ph) / 2.0
            else:
                # Second plate → centre of BOTTOM half
                ty = (a3_h / 2.0 - ph) / 2.0

            a3_page.merge_transformed_page(
                plate, Transformation().translate(tx=tx, ty=ty)
            )

        with open(output_path, 'wb') as f:
            writer.write(f)

        print(f"[MERGE] A3 PDF composed: {os.path.basename(output_path)}")
        return True

    except Exception as e:
        print(f"[MERGE] pypdf compose failed: {e}")
        traceback.print_exc()
        # Fallback: copy first individual PDF as the output
        try:
            shutil.copy2(input_paths[0], output_path)
            print(f"[MERGE] Fallback: copied first plate PDF as output.")
        except Exception as fb_e:
            print(f"[MERGE] Fallback also failed: {fb_e}")
        return False


def execute_print_merge_to_pdf(corel_app, data_records, output_pdf_path,
                               template_mv_path, template_mc_path,
                               global_dx=0.0, global_dy=0.0):
    """
    FINAL ALGORITHM (matches user specification):

    Step 1: For each record
            a. shutil.copy2  template → temp CDR  (original never modified)
            b. OpenDocument(temp CDR)
            c. Inject text IN-MEMORY (no cross-doc paste)
            d. Export PREVIEW + PRINT PDFs from the SAME document
            e. Close temp CDR
            f. Delete temp CDR

    Step 2: Compose individual plate PDFs onto one A3 page
            – 1 plate  → centred
            – 2 plates → top half / bottom half

    Step 3: Delete individual temp PDFs

    Step 4: Return.  Flutter waits for confirmation before printing.
    """
    try:
        total_records = len(data_records)
        print(f"[STAGE] Initializing Master Engine. Records: {total_records}")
        _heartbeat("Engine started")

        if total_records == 0:
            print("No records found.")
            return False

        base         = output_pdf_path.replace(".pdf", "")
        preview_path = f"{base}_PREVIEW.pdf"
        print_path   = f"{base}_PRINT.pdf"

        temp_previews = []
        temp_prints   = []

        # ── Step 1: Export each plate individually ────────────────────────────────
        for p_idx, record_data in enumerate(data_records):
            p_type        = record_data.get("type", "MV").upper()
            template_path = template_mc_path if p_type == "MC" else template_mv_path

            if not os.path.exists(template_path):
                print(f"Error: Template not found for record {p_idx+1}: {template_path}")
                continue

            temp_cdr  = template_path.replace('.cdr', f'__tmp_{p_idx}.cdr')
            t_preview = f"{base}_tmp{p_idx}_PREVIEW.pdf"
            t_print   = f"{base}_tmp{p_idx}_PRINT.pdf"

            print(f"\n[STEP {p_idx+1}/{total_records}] Processing {p_type} plate...")
            shutil.copy2(template_path, temp_cdr)
            print(f"  [COPY] {os.path.basename(temp_cdr)}")

            ok = _export_record_pdfs(
                corel_app, temp_cdr, record_data, p_type,
                t_preview, t_print
            )

            # Always remove temp CDR regardless of success
            try:
                os.remove(temp_cdr)
            except Exception as rm_e:
                print(f"  Warning: could not remove temp CDR: {rm_e}")

            if ok:
                temp_previews.append(t_preview)
                temp_prints.append(t_print)
            else:
                print(f"  Error: Export failed for record {p_idx+1}.")

            _heartbeat(f"Record {p_idx+1}/{total_records} exported")

        if not temp_previews:
            print("Error: No records exported successfully.")
            return False

        # ── Step 2: Compose individual PDFs onto one A3 page ─────────────────────
        print(f"\n[STAGE] Composing {len(temp_previews)} plate(s) onto A3 canvas...")
        _heartbeat("Starting A3 compose")

        _compose_a3(temp_previews, preview_path)
        _compose_a3(temp_prints,   print_path)
        _heartbeat("A3 compose complete")

        # ── Step 3: Cleanup temp PDFs ─────────────────────────────────────────────
        for p in temp_previews + temp_prints:
            try:
                os.remove(p)
            except:
                pass

        n = len(temp_previews)
        print(f"\nSilenced PDF exports achieved. {n}/{total_records} records exported.")
        return True

    except Exception as e:
        print(f"Print Merge Error: {e}")
        traceback.print_exc()
        return False
