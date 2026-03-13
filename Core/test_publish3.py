import win32com.client
import traceback
import time

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr"
doc = corel.OpenDocument(template_path)
doc.Unit = 11
page1 = doc.Pages.Item(1)

try:
    print("Duplicating...")
    sr = corel.CreateShapeRange()
    for i in range(1, page1.Shapes.Count + 1):
        s = page1.Shapes.Item(i)
        if s.Type != 9:
            sr.Add(s)
    sr.Copy()
    
    for target_idx in range(2, 6):
        doc.AddPages(1)
        doc.Pages.Item(target_idx).Activate()
        doc.ActiveLayer.Paste()

    print("SetSize...")
    for p_idx in range(1, 6):
        curr_page = doc.Pages.Item(p_idx)
        curr_page.SetSize(23.5, 13.5)
        
    corel.ActiveDocument.ClearSelection()
    
    doc.PDFSettings.PublishRange = 0
    out_pdf = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\test_out.pdf"
    print("Publishing to PDF...")
    doc.PublishToPDF(out_pdf)
    print("PDF export done.")
    
except Exception as e:
    traceback.print_exc()

doc.Dirty = False
doc.Close()
