import win32com.client
import traceback
import time

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr"
doc = corel.OpenDocument(template_path)

try:
    print("Testing mock merge and PDF export on MC...")
    page1 = doc.Pages.Item(1)
    
    # Just duplicate a page and paste something
    doc.AddPages(1)
    page1.Activate()
    sr = corel.CreateShapeRange()
    for i in range(1, page1.Shapes.Count + 1):
        s = page1.Shapes.Item(i)
        if s.Type != 9:
            sr.Add(s)
    sr.Copy()
    
    doc.Pages.Item(2).Activate()
    doc.ActiveLayer.Paste()
    
    print("Clearing selection...")
    doc.ClearSelection()
    
    doc.PDFSettings.PublishRange = 0
    
    out_pdf = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\test_out.pdf"
    print("Publishing to PDF...")
    doc.PublishToPDF(out_pdf)
    print("PDF export done.")
    
except Exception as e:
    traceback.print_exc()

doc.Dirty = False
doc.Close()
