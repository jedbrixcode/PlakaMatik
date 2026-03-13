import win32com.client
import traceback

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\MV_PLATE.cdr"
doc = corel.OpenDocument(template_path)

try:
    print("Testing copy with doc.Unit = 11...")
    doc.Unit = 11
    page1 = doc.Pages.Item(1)
    
    # Add new page
    page2 = doc.AddPages(1)
    
    # Copy shapes
    page1.Shapes.All().Copy()
    print("Copy successful.")
    
except Exception as e:
    traceback.print_exc()

doc.Dirty = False
doc.Close()
