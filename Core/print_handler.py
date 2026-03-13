import os
import time
import traceback

def execute_print_merge_to_pdf(corel_app, template_doc, data_records, output_pdf_path, plate_type="MV"):
    try:
        print(f"Initializing manual Python merge sequence... Template Type: {plate_type}")

        # Set unit to cm (11 = cdrCentimeter)
        template_doc.Unit = 11
        
        page1 = template_doc.Pages.Item(1)
        total_records = len(data_records)
        print(f"Total records to process: {total_records}")

        if total_records == 0:
            print("No records found to merge.")
            return False

        print("Duplicating template pages for each record...")
        
        # Step 1: Duplicate Page 1 shapes for all required records
        if total_records > 1:
            page1.Activate()
            
            # Create a ShapeRange of all non-guideline shapes to prevent COM clipboard crashes
            sr = corel_app.CreateShapeRange()
            for i in range(1, page1.Shapes.Count + 1):
                s = page1.Shapes.Item(i)
                if s.Type != 9: # 9 is cdrGuidelineShape
                    sr.Add(s)
            
            sr.Copy()
            
            for target_idx in range(2, total_records + 1):
                template_doc.AddPages(1)
                template_doc.Pages.Item(target_idx).Activate()
                template_doc.ActiveLayer.Paste()

        print("Pages duplicated. Mapping data to text placeholders...")

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
                            if "MIDDLE" in current_text:
                                s.Text.Story.Text = record.get("middle", "")
                except Exception as e:
                    pass
                
                try:
                    if s.Type == 7: # cdrGroupShape
                        replace_text_in_shapes(s.Shapes, record, p_type)
                    elif s.PowerClip:
                        replace_text_in_shapes(s.PowerClip.Shapes, record, p_type)
                except:
                    pass

        # Step 2: Apply the record data to each respective page
        for p_idx in range(1, total_records + 1):
            curr_page = template_doc.Pages.Item(p_idx)
            record_data = data_records[p_idx - 1]
            replace_text_in_shapes(curr_page.Shapes, record_data, plate_type)
            
            # Step 2b: Apply the size scaling NOW, after all shapes are placed
            if plate_type.upper() == "MV":
                curr_page.SetSize(39.0, 14.0)
            elif plate_type.upper() == "MC":
                curr_page.SetSize(23.5, 13.5)

        print(f"Data applied successfully. Exporting to {output_pdf_path}")

        # Deselect all to prevent export engine crash
        corel_app.ActiveDocument.ClearSelection()

        # Step 3: Publish to PDF
        pdf_settings = template_doc.PDFSettings
        pdf_settings.PublishRange = 0 # 0 means export whole docu
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
