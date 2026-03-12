import win32com.client
import os

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\MV_PLATE.cdr"
if not os.path.exists(template_path):
    print("Template not found!")
    exit()

# Create a temporary copy to avoid messing up the template
temp_cdr = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\Core\temp_merge.cdr"
import shutil
shutil.copyfile(template_path, temp_cdr)

doc = corel.OpenDocument(temp_cdr)
print("Doc opened.")

# Duplicate page 1
print("Duplicating page...")
# Page Duplicate method usually inserts after. Let's try doc.Pages(1).Duplicate()
page1 = doc.Pages.Item(1)
page2 = doc.InsertPagesEx(1, False, page1.Index, page1.SizeWidth, page1.SizeHeight) # Not duplicate?
# wait, page.Duplicate() might be easier if it exists, or doc.Pages(1).Shapes.All().Copy() and Paste()
try:
    # Try page.Duplicate()
    # Or doc.Pages(1).Layers("Layer 1").Shapes.All().Copy()
    page1 = doc.Pages.Item(1)
    new_page = doc.AddPages(1)
    new_page.SizeWidth = page1.SizeWidth
    new_page.SizeHeight = page1.SizeHeight
    
    # Copy all shapes from page 1 to new page
    # In VBA: ActiveDocument.Pages(1).Shapes.All().Copy
    page1.Shapes.All().Copy()
    
    # Paste
    doc.Pages.Item(2).Activate()
    doc.ActiveLayer.Paste()
    print("Pasted shapes to page 2.")
    
except Exception as e:
    print(f"Error duplicating: {e}")

doc.Close()
