import win32com.client
import os

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\MV_PLATE.cdr"
doc = corel.OpenDocument(template_path)
doc.Unit = 4 # cm

try:
    page1 = doc.Pages.Item(1)
    
    print(f"Original Size: {page1.SizeWidth} x {page1.SizeHeight}")
    
    # Let's get the first shape's original position
    s1 = page1.Shapes.Item(1)
    print(f"Shape 1 Original POS: X={s1.PositionX}, Y={s1.PositionY}")
    
    page1.SetSize(39.0, 14.0)
    
    print(f"Shape 1 New POS: X={s1.PositionX}, Y={s1.PositionY}")
    
    # The bounding box of the whole page content:
    sr = corel.CreateShapeRange()
    for i in range(1, page1.Shapes.Count + 1):
        sr.Add(page1.Shapes.Item(i))
    print(f"Group Bounding Box: Width={sr.SizeWidth}, Height={sr.SizeHeight}")

except Exception as e:
    import traceback
    traceback.print_exc()

doc.Dirty = False
doc.Close()
