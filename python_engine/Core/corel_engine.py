import win32com.client
import pyautogui
import time
import os

class CorelAutomator:
    def __init__(self):
        self.corel = None
        self.doc = None
        # Store the bypass delay so connect() can apply it at the right moment
        self._trial_delay = 0

    def bypass_trial_screen(self, delay=5):
        """
        Store the desired trial-screen wait delay.
        The actual bypass is executed inside connect() AFTER CorelDRAW
        has been dispatched and has had time to show its trial dialog.
        """
        self._trial_delay = delay
        if delay <= 0:
            print("Trial screen bypass disabled (delay=0). Will skip on connect.")
        else:
            print(f"Trial bypass armed: will wait {delay}s after CorelDRAW launches.")

    def connect(self):
        try:
            print("attempting to connect to CorelDRAW 2018...")
            # This call triggers CorelDRAW to launch (or attach to a running instance).
            self.corel = win32com.client.Dispatch("CorelDRAW.Application")
            self.corel.Visible = True

            # --- Trial screen bypass window ---
            # Now that CorelDRAW is launching, wait for the trial dialog to appear.
            if self._trial_delay > 0:
                print(f"Waiting {self._trial_delay}s for trial screen to appear...")
                time.sleep(self._trial_delay)
                pyautogui.hotkey('alt', 'z')
                print("Trial screen bypassed.")
            else:
                print("Trial bypass skipped (delay=0).")

            # Wait for CorelDRAW UI to fully stabilize after the bypass
            print("waiting for UI to stabilize...")
            time.sleep(7)

            # Set window to normal state
            try:
                self.corel.Frame.WindowState = 1
            except Exception as ui_error:
                print(f"Note: could not set windowstate (UI loading): {ui_error}")

            print("success! corelDraw connection established")
            return True

        except Exception as e:
            print("Failed to connect to CorelDRAW")
            print(f"Error details: {e}")
            return False

    def open_template(self, template_path):
        try:
            if not os.path.exists(template_path):
                print(f"Error: Template not found at {template_path}")
                return

            print(f"Opening LTO template: {template_path}")
            self.doc = self.corel.OpenDocument(template_path)
            print("Template successfully loaded")
        except Exception as e:
            print(f"Failed to open template: {e}")