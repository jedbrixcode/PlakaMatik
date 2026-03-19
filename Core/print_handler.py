import os
import time
import traceback

def execute_print_merge_to_pdf(corel_app, template_doc, data_records, output_pdf_path, plate_type="MV"):
    try:
        print(f"Initializing manual Python merge sequence... Template Type: {plate_type}")

        # Set unit to cm (4 = cdrCentimeter)
        template_doc.Unit = 4
        
        page1 = template_doc.Pages.Item(1)
        total_records = len(data_records)
        print(f"Total records to process: {total_records}")

        # check if any data was parsed
        if total_records == 0:
            print("No records found to merge.")
            return False

        print("Duplicating template pages for each record...")
        
        # Step 1: Duplicate Page 1 shapes for all required records
        if total_records > 1:
            page1.Activate()
            
            # Create a ShapeRange of all non-guideline shapes to prevent COM clipboard crashes
            sr = corel_app.CreateShapeRange()

            # iterate through each shape on page 1
            for i in range(1, page1.Shapes.Count + 1):
                s = page1.Shapes.Item(i)
                if s.Type != 9: # 9 is cdrGuidelineShape
                    sr.Add(s)
            
            sr.Copy()
            
            # duplicate page 1 for each record
            for target_idx in range(2, total_records + 1):
                template_doc.AddPages(1)
                template_doc.Pages.Item(target_idx).Activate()
                template_doc.ActiveLayer.Paste()

        print("Pages duplicated. Mapping data to text placeholders...")

        # Recursive text substitution function
        def replace_text_in_shapes(shapes, record, p_type):
            # iterate through each shape on page 1
            for i in range(1, shapes.Count + 1):
                s = shapes.Item(i)
                try:
                    if hasattr(s, 'Text') and s.Text:
                        current_text = s.Text.Story.Text
                        
                        # check plate type and replace text for MV plates
                        if p_type.  upper() == "MV":
                            if "MIDDLE" in current_text or "<MIDDLE>" in current_text:
                                s.Text.Story.Text = record.get("middle", "")
                            elif "IDENTIFIER" in current_text or "<IDENTIFIER>" in current_text:
                                s.Text.Story.Text = record.get("identifier", "")
                        
                        # check plate type and replace text for MC plates
                        elif p_type.upper() == "MC":
                            if "MIDDLE" in current_text:
                                s.Text.Story.Text = record.get("middle", "")
                except Exception as e:
                    pass
                
                # check if shape is a group or powerclip
                try:    
                    if s.Type == 7: # cdrGroupShape
                        replace_text_in_shapes(s.Shapes, record, p_type)
                    elif s.PowerClip:
                        replace_text_in_shapes(s.PowerClip.Shapes, record, p_type)
                except:
                    pass

        # Step 2: Apply the record data to each respective page
        for p_idx in range(1, total_records + 1):
            # get the current page and its record data
            curr_page = template_doc.Pages.Item(p_idx)
            record_data = data_records[p_idx - 1]   
            
            # replace the text in the current page
            replace_text_in_shapes(curr_page.Shapes, record_data, plate_type)
            
            # Record original sizes to calculate shift delta
            old_width = curr_page.SizeWidth
            old_height = curr_page.SizeHeight
            
            # Step 2b: Apply the size scaling NOW, after all shapes are placed
            new_width = 39.2
            new_height = 14.2
            if plate_type.upper() == "MV":
                new_width = 39.2
                new_height = 14.2
            elif plate_type.upper() == "MC":
                new_width = 24.0
                new_height = 14.0
                
            curr_page.SetSize(new_width, new_height)

            # Step 2c: Translate the template graphics dynamically to the new cropped center
            # By calculating the fixed offset of the page center shifting relative to bottom-left 0,0
            dx = (new_width / 1.825) - (old_width / 2.0)
            dy = (new_height / 2.1) - (old_height / 2.0)
            
            try:
                sr = corel_app.CreateShapeRange()
                for i in range(1, curr_page.Shapes.Count + 1):
                    s = curr_page.Shapes.Item(i)
                    if s.Type != 9: # Skip guidelines
                        sr.Add(s)
                
                if sr.Count > 0:
                    sr.Move(dx, dy)
            except Exception as center_ex:
                print(f"Warning: Could not align shapes on page {p_idx}: {center_ex}")

        print(f"Data applied successfully. Exporting to {output_pdf_path}")

        # Deselect all to prevent export engine crash
        corel_app.ActiveDocument.ClearSelection()

        # Step 3: Publish to PDF
        pdf_settings = template_doc.PDFSettings
        pdf_settings.PublishRange = 0 # 0 means export whole doc
        template_doc.PublishToPDF(output_pdf_path)
        print("PDF export done.")

        # Step 4: Close without saving to protect the template
        template_doc.Dirty = False
        template_doc.Close()
        print("temporary merge workplace cleared.")

        return True

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Print Merge Error: {e}")
        try:
            template_doc.Dirty = False
            template_doc.Close()
        except:
            pass
        return False
