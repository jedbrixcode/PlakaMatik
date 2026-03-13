import win32com.client
import traceback
import os

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\MV_PLATE.cdr"
doc = corel.OpenDocument(template_path)
doc.Unit = 11 # cm

try:
    print("Testing copy THEN size...")
    page1 = doc.Pages.Item(1)
    
    # 1. Add new page
    page2 = doc.AddPages(1)
    
    # 2. Copy shapes
    page1.Shapes.All().Copy()
    
    # 3. Paste
    doc.Pages.Item(2).Activate()
    doc.ActiveLayer.Paste()
    print("Copy and Paste successful.")
    
    # 4. Set sizes at the end
    print("Testing SetSize on both pages...")
    doc.Pages.Item(1).SetSize(39.0, 14.0)
    doc.Pages.Item(2).SetSize(39.0, 14.0)
    print("SetSize successful.")
    
except Exception as e:
    traceback.print_exc()

doc.Close(False)
