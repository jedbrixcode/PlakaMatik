import win32com.client
import traceback

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr"
doc = corel.OpenDocument(template_path)

try:
    print("Testing filtered copy on MC...")
    page1 = doc.Pages.Item(1)
    page1.Activate()
    
    # Create empty shape range
    # In CorelDRAW: Application.CreateShapeRange()
    sr = corel.CreateShapeRange()
    for i in range(1, page1.Shapes.Count + 1):
        s = page1.Shapes.Item(i)
        if s.Type != 9: # skip guidelines
            sr.Add(s)
            
    print(f"Added {sr.Count} shapes to range.")
    sr.Copy()
    print("Filtered copy successful.")
    
    # Try duplicating page as alternative
    print("Testing page duplication...")
    # page1.Duplicate() doesn't exist? Let's try it:
    # new_page = doc.InsertPagesEx(1, False, 1, 8.5, 11) -> we did this before, it didn't copy shapes.
    # What about doc.DuplicatePage(1)? No.
    
except Exception as e:
    traceback.print_exc()

doc.Dirty = False
doc.Close()
