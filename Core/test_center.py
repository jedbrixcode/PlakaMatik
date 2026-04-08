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
    
    # Let's see original size
    print(f"Original Size: {page1.SizeWidth} x {page1.SizeHeight}")
    
    # Create shape range of everything
    sr = corel.CreateShapeRange()
    for i in range(1, page1.Shapes.Count + 1):
        sr.Add(page1.Shapes.Item(i))
    
    # 1. Resize
    print("Resizing to 39 x 14...")
    page1.SetSize(39.0, 14.0)
    
    # 2. Try ShapeRange.Group().AlignToPageCenter(3) then Ungroup?
    # Actually, CorelDRAW has ShapeRange.PositionX and PositionY.
    # Grouping is easiest if we don't want to mess up.
    print("Grouping and centering...")
    grp = sr.Group()
    
    # cdrAlignHCenter = 1, cdrAlignVCenter = 2 -> 3
    # Wait, Group.AlignToPageCenter(3)
    grp.AlignToPageCenter(1)
    grp.AlignToPageCenter(2)
    
    # Ungroup it back so everything is exactly as it was
    grp.Ungroup()
    
    out_pdf = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\test_center_mv.pdf"
    doc.PDFSettings.PublishRange = 1 # current page
    doc.PublishToPDF(out_pdf)
    print("Exported to test_center_mv.pdf")
    
except Exception as e:
    import traceback
    traceback.print_exc()

doc.Dirty = False
doc.Close()
