import win32com.client
import traceback

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr"
doc = corel.OpenDocument(template_path)

try:
    print("Testing copy on MC...")
    page1 = doc.Pages.Item(1)
    print("Shapes count:", page1.Shapes.Count)
    
    # Add new page
    print("Adding page...")
    doc.AddPages(1)
    
    print("Selecting everything...")
    page1.Activate()
    sr = page1.Shapes.All()
    print("Selected shape range size:", sr.Count)
    
    print("Performing Copy...")
    sr.Copy()
    print("Copy successful.")
    
except Exception as e:
    traceback.print_exc()

doc.Dirty = False
doc.Close()
