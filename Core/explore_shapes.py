import win32com.client
import os

def explore_shapes(shapes, parent_name=""):
    for i in range(1, shapes.Count + 1):
        s = shapes.Item(i)
        
        try:
            if hasattr(s, 'Text') and s.Text:
                print(f"[{s.Type}] {parent_name}/{s.Name}")
                print(f"   -> TEXT: {s.Text.Story.Text}")
            else:
                # Still print if it's a group
                if s.Type == 7 or s.Type == 6:
                    print(f"[{s.Type}] {parent_name}/{s.Name}")
        except Exception as e:
            pass
            
        try:
            if s.Type == 7: # group
                explore_shapes(s.Shapes, parent_name + "/" + s.Name)
            elif s.PowerClip:
                explore_shapes(s.PowerClip.Shapes, parent_name + "(PowerClip)")
        except:
            pass

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_paths = [
    r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\MV_PLATE.cdr",
    r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr"
]

for template_path in template_paths:
    print(f"\n--- Checking {os.path.basename(template_path)} ---")
    if not os.path.exists(template_path):
        print("Template not found!")
        continue

    doc = corel.OpenDocument(template_path)
    
    for p in doc.Pages:
        for l in p.Layers:
            explore_shapes(l.Shapes, f"Page{p.Index}/Layer_{l.Name}")

    doc.Close()

print("Done exploring.")
