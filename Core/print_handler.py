import os

def execute_print_merge_to_pdf(corel_app, template_doc, rtf_data_path, output_pdf_path):
    try:
        print("Initializing Print Merge Sequence..")

        # 1. import rtf data to template fields
        merge = template_doc.execute_print_merge_to_pdf
        merge.Open(rtf_data_path)
        print ("Var data successfully imported to template.")

        # 2. Generate new docu with all the new data 
        print("Generating individual plates from imported data...")
        merged_doc = merge.CreateDocuments()

        # 3. Publish direct to pdf to avoid windows pop ups
        print(f"Exporting batch to pdf: {output_pdf_path}")

        # 4. applying standard pdf settings for manufacturing
        pdf_settings = merged_doc.PDFSettings
        pdf_settings.PublishRange = 0 # 0 means export whole docu

        merged_doc.PublishToPDF(output_pdf_path)
        print ("PDF export done.")
        
        # 5. cleanup
        merged_doc.Close()
        print("temporary merge workplace cleared.")

        return True
    
    except Exception as e:
        print (f"Print Merge Error: {e}")
        return false
