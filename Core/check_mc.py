import win32com.client
import os

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True
doc = corel.OpenDocument(r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr")

# Set units to cm
doc.Unit = 1

def explore_shapes(shapes):
    for i in range(1, shapes.Count + 1):
        s = shapes.Item(i)
        try:
            if hasattr(s, 'Text') and s.Text:
                y = s.PositionY
                print(f"[{s.Type}] {s.Name}")
                print(f"   -> TEXT: {s.Text.Story.Text}")
                print(f"   -> POS Y: {y} cm")
        except Exception as e:
            pass
            
        try:
            if s.Type == 7: # group
                explore_shapes(s.Shapes)
            elif s.PowerClip:
                explore_shapes(s.PowerClip.Shapes)
        except:
            pass

for p in doc.Pages:
    for l in p.Layers:
        explore_shapes(l.Shapes)

doc.Close()
