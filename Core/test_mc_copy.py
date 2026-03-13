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
    
    print("Looping through shapes...")
    page1.Activate()
    for i in range(1, page1.Shapes.Count + 1):
        s = page1.Shapes.Item(i)
        print(f"[{i}] Type: {s.Type}, Name: {s.Name}")
        try:
            s.Copy()
            print("  -> Copied successfully.")
        except Exception as e:
            print(f"  -> Failed: {e}")
            
    print("Trying to group them and copy...")
    
except Exception as e:
    traceback.print_exc()

doc.Dirty = False
doc.Close()
