import win32com.client
import os

print("Connecting to CorelDRAW...")
corel = win32com.client.Dispatch("CorelDRAW.Application")
corel.Visible = True

# Check if GMSManager exists
try:
    gms = getattr(corel, 'GMSManager', None)
    if gms:
        print("GMSManager found:", gms)
        print("Methods:", dir(gms))
    else:
        print("No GMSManager.")
        
    print("Application Methods:", [m for m in dir(corel) if 'macro' in m.lower() or 'gms' in m.lower() or 'vba' in m.lower()])
except Exception as e:
    print(e)
