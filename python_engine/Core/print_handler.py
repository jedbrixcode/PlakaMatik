import os
import time
import traceback

def execute_print_merge_to_pdf(corel_app, data_records, output_pdf_path, template_mv_path, template_mc_path):
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
        master_layer_payload = master_page.Layers.Item("Layer 1")
        master_layer_payload.Name = "Payload"
        
        master_layer_guides = master_page.CreateLayer("Guides")
        master_layer_guides.Printable = True

        # Recursive text substitution function
        def replace_text_in_shapes(shapes, record, p_type):
            for i in range(1, shapes.Count + 1):
                s = shapes.Item(i)
                try:
                    if hasattr(s, 'Text') and s.Text:
                        current_text = s.Text.Story.Text
                        
                        if p_type.upper() == "MV":
                            if "MIDDLE" in current_text or "<MIDDLE>" in current_text:
                                s.Text.Story.Text = record.get("middle", "")
                            elif "IDENTIFIER" in current_text or "<IDENTIFIER>" in current_text:
                                s.Text.Story.Text = record.get("identifier", "")
                        
                        elif p_type.upper() == "MC":
                            # Allow un-identifier'd templates to gracefully ignore identifier logic natively
                            if "MIDDLE" in current_text or "<MIDDLE>" in current_text:
                                s.Text.Story.Text = record.get("middle", "")
                            elif "IDENTIFIER" in current_text or "<IDENTIFIER>" in current_text:
                                s.Text.Story.Text = record.get("identifier", "")
                except Exception as e:
                    pass
                
                try:    
                    if s.Type == 7: # cdrGroupShape
                        replace_text_in_shapes(s.Shapes, record, p_type)
                    elif s.PowerClip:
                        replace_text_in_shapes(s.PowerClip.Shapes, record, p_type)
                except:
                    pass

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

            # Map Text values directly into the template
            replace_text_in_shapes(temp_page.Shapes, record_data, p_type)

            # -------------------------------------------------------------
            # LAYER ISOLATION LOGIC 
            # -------------------------------------------------------------
            # Extract Text Payload (Layer 1)
            sr_payload = corel_app.CreateShapeRange()
            try:
                layer_1 = temp_page.Layers.Item("Layer 1")
                for i in range(1, layer_1.Shapes.Count + 1):
                    s = layer_1.Shapes.Item(i)
                    if s.Type != 9:
                        sr_payload.Add(s)
            except Exception as le:
                    print(f"Warning: Could not fetch Layer 1 from template: {le}")

            dx = 0.0
            dy = 0.0

            if sr_payload.Count > 0:
                grouped_payload = sr_payload.Group()
                grouped_payload.Copy() 
                time.sleep(0.5) 

                master_doc.Activate()
                master_layer_payload.Activate()
                pasted_payload = master_layer_payload.Paste()
                time.sleep(0.5) 

                cx = pasted_payload.PositionX + (pasted_payload.SizeWidth / 2.0)
                cy = pasted_payload.PositionY - (pasted_payload.SizeHeight / 2.0)

                target_px = 42.0 / 2.0 
                if total_records == 1:
                    target_py = 29.7 / 2.0 
                else:
                    target_py = 29.7 * 0.75 if p_idx == 0 else 29.7 * 0.25 
                
                dx = target_px - cx
                dy = target_py - cy
                
                pasted_payload.Move(dx, dy)

            # Extract Visual Template Outlines (Guides)
            sr_guides = corel_app.CreateShapeRange()
            try:
                layer_guides = temp_page.Layers.Item("Guides")
                for i in range(1, layer_guides.Shapes.Count + 1):
                    s = layer_guides.Shapes.Item(i)
                    if s.Type != 9:
                        sr_guides.Add(s)
            except Exception as le:
                    print(f"Warning: Could not fetch Guides from template: {le}")
            
            if sr_guides.Count > 0:
                grouped_guides = sr_guides.Group()
                grouped_guides.Copy()
                time.sleep(0.5)
                
                master_doc.Activate()
                master_layer_guides.Activate()
                pasted_guides = master_layer_guides.Paste()
                time.sleep(0.5)
                
                # Lock visual guides mathematically to the exact same position as payload
                pasted_guides.Move(dx, dy)
                
            time.sleep(0.5)
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
        
        # EXPORT 1: The UI Preview PDF (Includes Visible Guides)
        print("Data mapping securely applied. Exporting VERIFICATION PREVIEW...")
        master_layer_guides.Printable = True
        master_layer_guides.Visible = True
        
        preview_pdf_path = output_pdf_path.replace(".pdf", "_PREVIEW.pdf")
        master_doc.PublishToPDF(preview_pdf_path)
        
        # EXPORT 2: The Physical UV Plate PDF (Naked Payload Only)
        print("Exporting PRINT-READY PAYLOAD...")
        master_layer_guides.Printable = False
        master_layer_guides.Visible = False
        
        print_pdf_path = output_pdf_path.replace(".pdf", "_PRINT.pdf")
        master_doc.PublishToPDF(print_pdf_path)

        print("Silenced PDF exports achieved flawlessly.")

        master_doc.Dirty = False
        master_doc.Close()
        print("Master workspace cleared perfectly.")

        return True

    except Exception as e:
        print(f"Print Merge Error: {e}")
        import traceback
        traceback.print_exc()
        try:
            if master_doc:
                master_doc.Dirty = False
                master_doc.Close()
        except:
            pass
        return False
