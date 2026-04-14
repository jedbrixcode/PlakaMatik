import os
import time
import traceback

def execute_print_merge_to_pdf(corel_app, data_records, output_pdf_path, template_mv_path, template_mc_path):
    try:
        total_records = len(data_records)
        print(f"Initializing Master Engine merge. Total records to process: {total_records}")

        # check if any data was parsed
        if total_records == 0:
            print("No records found to merge.")
            return False

        # Create Master A3 Document
        print("Creating Master A3 payload layout...")
        master_doc = corel_app.CreateDocument()
        master_doc.Unit = 4 # cdrCentimeter
        master_page = master_doc.Pages.Item(1)
        master_page.SetSize(42.0, 29.7) # Strictly A3 paper size

        # Recursive text substitution function
        def replace_text_in_shapes(shapes, record, p_type):
            for i in range(1, shapes.Count + 1):
                s = shapes.Item(i)
                try:
                    if hasattr(s, 'Text') and s.Text:
                        current_text = s.Text.Story.Text
                        
                        # Apply to MV
                        if p_type.upper() == "MV":
                            if "MIDDLE" in current_text or "<MIDDLE>" in current_text:
                                s.Text.Story.Text = record.get("middle", "")
                            elif "IDENTIFIER" in current_text or "<IDENTIFIER>" in current_text:
                                s.Text.Story.Text = record.get("identifier", "")
                        
                        # Apply to MC
                        elif p_type.upper() == "MC":
                            if "MIDDLE" in current_text or "<MIDDLE>" in current_text:
                                s.Text.Story.Text = record.get("middle", "")
                            elif "IDENTIFIER" in current_text or "<IDENTIFIER>" in current_text:
                                s.Text.Story.Text = record.get("identifier", "")
                except Exception as e:
                    pass
                
                # Recursively check grouped shapes and powerclips
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

            # Open target Template safely for extraction
            print(f"Opening template for Record {p_idx+1}/{total_records} Type: {p_type}")
            temp_doc = corel_app.OpenDocument(template_path)
            temp_doc.Unit = 4 # cdrCentimeter
            temp_page = temp_doc.Pages.Item(1)

            # Replace the text placeholders inside the template BEFORE bounding
            replace_text_in_shapes(temp_page.Shapes, record_data, p_type)

            # Group all shapes on page to mathematically lock relative structures
            sr = corel_app.CreateShapeRange()
            for i in range(1, temp_page.Shapes.Count + 1):
                s = temp_page.Shapes.Item(i)
                if s.Type != 9: # Skip guidelines
                    sr.Add(s)

            if sr.Count > 0:
                grouped_shape = sr.Group()
                grouped_shape.Copy() # Push to COM clipboard securely

                # Bring master document to foreground
                master_doc.Activate()
                # Paste shape onto A3 master layer
                pasted_shape = master_doc.ActiveLayer.Paste()

                # Calculate placement logic utilizing Algebraic centering mapping
                # Base constraints: Corel Y-axis starts bottom-left (0,0) going upwards.
                cx = pasted_shape.PositionX + (pasted_shape.SizeWidth / 2.0)
                cy = pasted_shape.PositionY - (pasted_shape.SizeHeight / 2.0)

                target_px = 42.0 / 2.0 # 21.0 - Dead center horizontal
                
                if p_idx == 0:
                    # Index 0 gets strictly aligned to the TOP half (Top Plate)
                    target_py = 29.7 * 0.75 
                else:
                    # Index 1 gets strictly aligned to the BOTTOM half (Bottom Plate)
                    target_py = 29.7 * 0.25 
                
                dx = target_px - cx
                dy = target_py - cy
                
                # Perform the transformation delta lock
                pasted_shape.Move(dx, dy)
            
            # Important: Close the original layout template without saving to keep it 100% pristine
            temp_doc.Dirty = False
            temp_doc.Close()

        # Step 3: Publish unified payload directly to PDF buffer
        print(f"Data mapping securely applied to A3 target. Exporting to {output_pdf_path}")
        master_doc.ClearSelection()
        
        pdf_settings = master_doc.PDFSettings
        pdf_settings.PublishRange = 0 # Export whole doc array silently
        master_doc.PublishToPDF(output_pdf_path)
        print("Silenced PDF export achieved.")

        # Step 4: Dismantle isolated workplace memory state
        master_doc.Dirty = False
        master_doc.Close()
        print("Master workspace cleared perfectly.")

        return True

    except Exception as e:
        print(f"Print Merge Error: {e}")
        import traceback
        traceback.print_exc()
        try:
            # Fallback failsafes
            if master_doc:
                master_doc.Dirty = False
                master_doc.Close()
        except:
            pass
        return False
