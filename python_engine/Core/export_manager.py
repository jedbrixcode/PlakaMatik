import os
import traceback
from text_mapper import replace_text_in_shapes
from layer_composer import compose_and_align_layers

def execute_print_merge_to_pdf(corel_app, data_records, output_pdf_path, template_mv_path, template_mc_path, global_dx=0.0, global_dy=0.0):
    try:
        total_records = len(data_records)
        print(f"Initializing Master Engine merge. Total records to process: {total_records}")

        if total_records == 0:
            print("No records found to merge.")
            return False

        # Create Master A3 Document
        print("Creating Master A3 payload layout...")
        master_doc = corel_app.CreateDocument()
        master_doc.Unit = 4 # cdrCentimeter
        master_page = master_doc.Pages.Item(1)
        master_page.SetSize(42.0, 29.7) # Strictly A3 paper size

        # Initialize Master Layers for separation
        master_layers = {
            'payload': master_page.Layers.Item("Layer 1"),
            'mv_guides': master_page.CreateLayer("MV_Guides"),
            'bg': master_page.CreateLayer("Background")
        }
        
        master_layers['payload'].Name = "Payload"
        master_layers['mv_guides'].Printable = True
        master_layers['bg'].Printable = False # Never printed, purely for UI previews

        # Step 2: Loop over the chunked records array
        for p_idx, record_data in enumerate(data_records):
            p_type = record_data.get("type", "MV").upper()
            template_path = template_mc_path if p_type == "MC" else template_mv_path

            if not os.path.exists(template_path):
                print(f"Error: Template {template_path} not found.")
                continue

            print(f"Opening template for Record {p_idx+1}/{total_records} Type: {p_type}")
            temp_doc = corel_app.OpenDocument(template_path)
            temp_doc.Unit = 4 # cdrCentimeter
            temp_page = temp_doc.Pages.Item(1)

            # Route recursive text values onto COM node structures
            replace_text_in_shapes(temp_page.Shapes, record_data, p_type)

            # Route geometric jig logic across COM layers
            compose_and_align_layers(corel_app, master_doc, master_layers, temp_page, p_type, p_idx, global_dx, global_dy)
            
            # Safely close without trashing COM queue
            try:
                temp_doc.Activate()
                temp_doc.Dirty = False
                temp_doc.Close()
            except Exception as inner_e:
                print(f"Warning: RPC suppressed during temp_doc closure {inner_e}")

        # Step 3: Dual PDF Export Strategy
        master_doc.ClearSelection()
        pdf_settings = master_doc.PDFSettings
        pdf_settings.PublishRange = 0 
        pdf_settings.ColorMode = 1 # Force RGB Native Rip Space
        
        # EXPORT 1: The UI Preview PDF (Includes Visible Backgrounds & Guides)
        print("Data mapping securely applied. Exporting VERIFICATION PREVIEW...")
        master_layers['bg'].Printable = True
        master_layers['bg'].Visible = True
        master_layers['mv_guides'].Printable = True
        master_layers['mv_guides'].Visible = True
        
        preview_pdf_path = output_pdf_path.replace(".pdf", "_PREVIEW.pdf")
        master_doc.PublishToPDF(preview_pdf_path)
        
        # EXPORT 2: The Physical UV Plate PDF (Naked Payload & Native Guides)
        print("Exporting PRINT-READY PAYLOAD...")
        master_layers['bg'].Printable = False
        master_layers['bg'].Visible = False
        master_layers['mv_guides'].Printable = True
        master_layers['mv_guides'].Visible = True
        
        print_pdf_path = output_pdf_path.replace(".pdf", "_PRINT.pdf")
        master_doc.PublishToPDF(print_pdf_path)

        print("Silenced PDF exports achieved flawlessly.")

        master_doc.Dirty = False
        master_doc.Close()
        print("Master workspace cleared perfectly.")

        return True

    except Exception as e:
        print(f"Print Merge Error: {e}")
        traceback.print_exc()
        try:
            if master_doc:
                master_doc.Dirty = False
                master_doc.Close()
        except:
            pass
        return False
