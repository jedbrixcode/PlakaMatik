import win32com.client
import os

def explore_shapes(shapes, parent_name=""):
    for i in range(1, shapes.Count + 1):
        s = shapes.Item(i)
        print(f"[{s.Type}] {parent_name}/{s.Name} ")
        
        if s.Type == 14: # text
            try:
                print(f"   -> TEXT: {s.Text.Story.Text}")
            except Exception as e:
                print(f"   -> TEXT ERROR: {e}")
        elif s.Type == 7: # group
            explore_shapes(s.Shapes, parent_name + "/" + s.Name)
        else:
            # check powerclip
            try:
                if s.PowerClip:
                    explore_shapes(s.PowerClip.Shapes, parent_name + "/" + s.Name + "(Powerclip)")
            except:
                pass
        
        # Are there any other properties we can check?
        try:
            if hasattr(s, 'Text') and s.Text:
                print(f"   -> Has Text property: {s.Text.Story.Text}")
        except: pass

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

template_path = r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\MV_PLATE.cdr"
doc = corel.OpenDocument(template_path)
print("Doc opened.")

print("--- REGULAR PAGES ---")
for p in doc.Pages:
    for l in p.Layers:
        explore_shapes(l.Shapes, f"Page{p.Index}/Layer_{l.Name}")

print("Done exploring.")
doc.Close()
