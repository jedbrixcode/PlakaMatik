import win32com.client
import pyautogui
import time
import os

class CorelAutomator:
    def __init__(self):
        self.corel = None
        self.doc = None

    # Trial screen bypass
    def bypass_trial_screen(self):
        time.sleep(1)
        print("waiting for the trial screen...")
        pyautogui.press('esc')
        print('close successfully clicked')

    # Connect to CorelDRAW
    def connect(self):
        try:
            # Connect to CorelDRAW
            print("attempting to connect to CorelDRAW 2018...")
            self.corel = win32com.client.Dispatch("CorelDRAW.Application")
            self.corel.Visible = True

            # Wait for UI to stabilize
            time.sleep(7)
            print("waiting for UI to stabilize...")

            # set window state to normal
            try: 
                self.corel.Frame.WindowState = 1
            except Exception as ui_error:
                print(f"Note: could not set windowstate (UI loading)") 

            print("success! corelDraw connection established")
            return True

        except Exception as e:
            print("Failed to connect to CorelDRAW")
            print(f"Error details: {e}")
            return False

    # open template specified by the operator
    def open_template(self, template_path):
        try:
            # check if template exists
            if not os.path.exists(template_path):
                print(f"Error: Template not found at {template_path}")
                return
            
            print(f"Opening LTO template: {template_path}")
            # OpenDocument is the correct COM method for existing files
            self.doc = self.corel.OpenDocument(template_path)
            print("Template successfully loaded")
        except Exception as e:
            print(f"Failed to open template: {e}")