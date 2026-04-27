import os
import traceback
from text_mapper import replace_text_in_shapes

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
    print(f"Warning: Layer '{layer_name_upper}' not found on page.")


def _export_single_record(corel_app, record_data, template_path, preview_pdf_path, print_pdf_path, global_dx=0.0, global_dy=0.0):
    """
    Opens a single CDR template, injects text values into 'Print Layer',
    then exports two PDFs: one PREVIEW (Print + Mock, no Guides) and one PRINT (Print only, no Mock/Guides).
    """
    p_type = record_data.get("type", "MV").upper()

    if not os.path.exists(template_path):
        print(f"Error: Template {template_path} not found.")
        return False

    print(f"Opening template: {os.path.basename(template_path)} for type {p_type}")
    doc = corel_app.OpenDocument(template_path)
    doc.Unit = 4  # cdrCentimeter
    page = doc.Pages.Item(1)

    # --- Inject text values into PRINT LAYER shapes only ---
    print(f"Injecting values: middle='{record_data.get('middle','')}' identifier='{record_data.get('identifier','')}' into Print Layer...")
    for i in range(1, page.Layers.Count + 1):
        lyr = page.Layers.Item(i)
        if lyr.Name.upper() == "PRINT LAYER":
            replace_text_in_shapes(lyr.Shapes, record_data, p_type)
            break

    doc.ClearSelection()
    pdf_settings = doc.PDFSettings
    pdf_settings.PublishRange = 0
    pdf_settings.ColorMode = 1  # Force RGB

    # --- EXPORT 1: PREVIEW PDF (Print Layer + Mock Layer visible, Guides hidden) ---
    print("Exporting VERIFICATION PREVIEW...")
    _set_layer_visibility(page, "PRINT LAYER", printable=True,  visible=True)
    _set_layer_visibility(page, "MOCK LAYER",  printable=True,  visible=True)
    _set_layer_visibility(page, "GUIDES",       printable=False, visible=False)
    doc.PublishToPDF(preview_pdf_path)
    print(f"Preview PDF saved: {preview_pdf_path}")

    # --- EXPORT 2: PRINT PDF (Print Layer only, Mock + Guides hidden) ---
    print("Exporting PRINT-READY PAYLOAD...")
    _set_layer_visibility(page, "PRINT LAYER", printable=True,  visible=True)
    _set_layer_visibility(page, "MOCK LAYER",  printable=False, visible=False)
    _set_layer_visibility(page, "GUIDES",       printable=False, visible=False)
    doc.PublishToPDF(print_pdf_path)
    print(f"Print PDF saved: {print_pdf_path}")

    doc.Dirty = False
    doc.Close()
    return True


def execute_print_merge_to_pdf(corel_app, data_records, output_pdf_path, template_mv_path, template_mc_path, global_dx=0.0, global_dy=0.0):
    """
    Entry point. For each record, opens the correct CDR template, injects text,
    and exports PREVIEW and PRINT PDFs directly from the template document.
    All outputs land in Documents/PlakaMatik Files/Outputs/.
    """
    try:
        total_records = len(data_records)
        print(f"Initializing Master Engine merge. Total records to process: {total_records}")

        if total_records == 0:
            print("No records found to merge.")
            return False

        base = output_pdf_path.replace(".pdf", "")
        preview_path = f"{base}_PREVIEW.pdf"
        print_path   = f"{base}_PRINT.pdf"

        if total_records == 1:
            # ---- SINGLE PLATE: open template, inject, export directly ----
            record_data = data_records[0]
            p_type = record_data.get("type", "MV").upper()
            template_path = template_mc_path if p_type == "MC" else template_mv_path
            result = _export_single_record(
                corel_app, record_data, template_path,
                preview_path, print_path, global_dx, global_dy
            )
            print(f"Silenced PDF exports achieved. {'1' if result else '0'}/1 records exported.")
            return result

        else:
            # ---- BATCH: combine all plates into one document, one PDF each ----
            print(f"Batch mode: combining {total_records} plates into a single A3 document...")

            # Open the first template to use as the master document
            first_record = data_records[0]
            first_type   = first_record.get("type", "MV").upper()
            first_tmpl   = template_mc_path if first_type == "MC" else template_mv_path

            if not os.path.exists(first_tmpl):
                print(f"Error: Template {first_tmpl} not found.")
                return False

            print(f"Opening master template: {os.path.basename(first_tmpl)}")
            master_doc = corel_app.OpenDocument(first_tmpl)
            master_doc.Unit = 4
            master_page1 = master_doc.Pages.Item(1)

            # Inject values into page 1
            print(f"Injecting record 1 ({first_type}) values into Page 1...")
            for i in range(1, master_page1.Layers.Count + 1):
                lyr = master_page1.Layers.Item(i)
                if lyr.Name.upper() == "PRINT LAYER":
                    from text_mapper import replace_text_in_shapes
                    replace_text_in_shapes(lyr.Shapes, first_record, first_type)
                    break

            # Add remaining records as additional pages
            for p_idx in range(1, total_records):
                record_data = data_records[p_idx]
                p_type = record_data.get("type", "MV").upper()
                template_path = template_mc_path if p_type == "MC" else template_mv_path

                if not os.path.exists(template_path):
                    print(f"Error: Template {template_path} not found. Skipping record {p_idx+1}.")
                    continue

                print(f"Opening template for record {p_idx+1}: {os.path.basename(template_path)}")
                src_doc = corel_app.OpenDocument(template_path)
                src_page = src_doc.Pages.Item(1)

                # Inject values
                print(f"Injecting record {p_idx+1} ({p_type}) values...")
                for i in range(1, src_page.Layers.Count + 1):
                    lyr = src_page.Layers.Item(i)
                    if lyr.Name.upper() == "PRINT LAYER":
                        from text_mapper import replace_text_in_shapes
                        replace_text_in_shapes(lyr.Shapes, record_data, p_type)
                        break

                # Add a new page to master document and copy all content
                try:
                    import time
                    new_page = master_doc.Pages.Add()
                    new_page.SetSize(src_page.SizeWidth, src_page.SizeHeight)

                    # Copy every layer's shapes from src_page → new_page
                    for li in range(1, src_page.Layers.Count + 1):
                        src_lyr = src_page.Layers.Item(li)
                        lyr_name = src_lyr.Name

                        # Skip guideline layers (Type 2)
                        if src_lyr.Shapes.Count == 0:
                            continue

                        # Select all shapes on this source layer
                        sr = corel_app.CreateShapeRange()
                        for si in range(1, src_lyr.Shapes.Count + 1):
                            try:
                                sr.Add(src_lyr.Shapes.Item(si))
                            except:
                                pass

                        if sr.Count == 0:
                            continue

                        # Group, copy, activate target page, paste
                        try:
                            grp = sr.Group()
                            grp.Copy()
                        except:
                            sr.Item(1).Copy()

                        time.sleep(0.3)
                        master_doc.Activate()

                        # Ensure matching layer exists on new_page
                        dest_lyr = None
                        try:
                            dest_lyr = new_page.Layers.Item(lyr_name)
                        except:
                            try:
                                dest_lyr = new_page.CreateLayer(lyr_name)
                            except:
                                dest_lyr = new_page.Layers.Item(1)
                        dest_lyr.Activate()

                        new_page.Activate()
                        master_doc.ActiveLayer.Paste()
                        time.sleep(0.3)

                    print(f"  Record {p_idx+1} merged into master doc as Page {master_doc.Pages.Count}.")
                except Exception as dup_e:
                    print(f"Warning: Page merge error for record {p_idx+1}: {dup_e}")

                src_doc.Dirty = False
                src_doc.Close()


            # Export PREVIEW: all layers visible except Guides
            print("Exporting COMBINED VERIFICATION PREVIEW...")
            for pg_idx in range(1, master_doc.Pages.Count + 1):
                pg = master_doc.Pages.Item(pg_idx)
                _set_layer_visibility(pg, "PRINT LAYER", printable=True,  visible=True)
                _set_layer_visibility(pg, "MOCK LAYER",  printable=True,  visible=True)
                _set_layer_visibility(pg, "GUIDES",       printable=False, visible=False)

            master_doc.ClearSelection()
            pdf_settings = master_doc.PDFSettings
            pdf_settings.PublishRange = 0
            pdf_settings.ColorMode = 1
            master_doc.PublishToPDF(preview_path)
            print(f"Combined preview PDF saved: {preview_path}")

            # Export PRINT: Print Layer only, Mock + Guides hidden
            print("Exporting COMBINED PRINT-READY PAYLOAD...")
            for pg_idx in range(1, master_doc.Pages.Count + 1):
                pg = master_doc.Pages.Item(pg_idx)
                _set_layer_visibility(pg, "PRINT LAYER", printable=True,  visible=True)
                _set_layer_visibility(pg, "MOCK LAYER",  printable=False, visible=False)
                _set_layer_visibility(pg, "GUIDES",       printable=False, visible=False)

            master_doc.PublishToPDF(print_path)
            print(f"Combined print PDF saved: {print_path}")

            master_doc.Dirty = False
            master_doc.Close()
            print(f"Silenced PDF exports achieved. {total_records}/{total_records} records exported.")
            return True

    except Exception as e:
        print(f"Print Merge Error: {e}")
        traceback.print_exc()
        return False
