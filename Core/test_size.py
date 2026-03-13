import win32com.client
import os

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

templates = [
    r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\MV_PLATE.cdr",
    r"c:\Users\Window 10\Documents\Jed Internship\Project\Plate Manufacturing Layout maker\CorelDRAW Templates\Protocol Plates MC.cdr"
]

for t in templates:
    doc = corel.OpenDocument(t)
    doc.Unit = 11 # cm
    p1 = doc.Pages.Item(1)
    print(f"Template: {os.path.basename(t)}")
    print(f"  Width: {p1.SizeWidth} cm, Height: {p1.SizeHeight} cm")
    doc.Close()
