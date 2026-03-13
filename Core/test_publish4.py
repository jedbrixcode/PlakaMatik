import win32com.client
import traceback
import time

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr"
doc = corel.OpenDocument(template_path)
doc.Unit = 11

def replace_text_in_shapes(shapes):
    for i in range(1, shapes.Count + 1):
        s = shapes.Item(i)
        try:
            if hasattr(s, 'Text') and s.Text:
                current_text = s.Text.Story.Text
                if "MIDDLE" in current_text or "HYBRID" in current_text:
                    s.Text.Story.Text = "TEST"
        except Exception as e:
            pass
        try:
            if s.Type == 7: # group
                replace_text_in_shapes(s.Shapes)
            elif s.PowerClip:
                replace_text_in_shapes(s.PowerClip.Shapes)
        except:
            pass

try:
    print("Testing mock merge and PDF export on MC...")
    page1 = doc.Pages.Item(1)
    
    # 1. Add pages and copy
    sr = corel.CreateShapeRange()
    for i in range(1, page1.Shapes.Count + 1):
        s = page1.Shapes.Item(i)
        if s.Type != 9:
            sr.Add(s)
    sr.Copy()
    
    for target_idx in range(2, 6): # 5 pages total
        doc.AddPages(1)
        doc.Pages.Item(target_idx).Activate()
        doc.ActiveLayer.Paste()
        
    # 2. Mutate text + SetSize
    for p_idx in range(1, 6):
        curr_page = doc.Pages.Item(p_idx)
        replace_text_in_shapes(curr_page.Shapes)
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
