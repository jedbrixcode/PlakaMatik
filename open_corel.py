import os
import win32com.client
import pyautogui
import time

# --- defined values ---
# this is where the script is saved
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Input: output from flutter based on user input with defined 4 spaces
input_txt_path = os.path.join(BASE_DIR, "temp_user_input.txt")

# output: rtf file for CorelDRAW Print Merge
output_rtf_path = os.path.join(BASE_DIR, "lto_print_merge.rtf")

def export_data_to_rtf():
    # Standard header for CorelDRAW files
    rtf_header = r"{\rtf\ansi\ud\uc1 2\par "
    rtf_footer = r"}"

    # storing formatted rows
    formatted_rows = []

    # Headers for the LTO plate template variables
    headers = r"\\MIDDLE\\\\IDENTIFIER\\\par "
    formatted_rows.append(headers)

    try:
        # checking if the flutter output exists
        if not os.path.exists(input_txt_path):
            print(f"crit error: {input_txt_path} not found")
            return

        with open(input_txt_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            print(f"Reading {len(lines)} lines from input...")

            # loop through raw text data
            for line in lines:
                # remove newline but keep the defined 4 spaces in the middle
                clean_line = line.rstrip()
                
                # skip the header row or empty lines if they exist
                if not clean_line or "Variable 1" in clean_line:
                    continue

                # split the line by the defined 4 spaces from Flutter
                # this ensures "20TH CONGRESS" stays as one variable
                columns = clean_line.split('    ')

                # if the row has data in both columns, formatting for corelDRAW
                if len(columns) >= 2:
                    var1 = columns[0].strip()
                    var2 = columns[1].strip()

                    # wrap data in corelDRAW in rtf tags
                    row_string = f"\\\\{var1}\\\\\\\\{var2}\\\\\\par "
                    formatted_rows.append(row_string)
                    print(f"Data parsed: {var1} | {var2}")
                else:
                    print(f"Skipping line (missing 4-space delimiter): {clean_line}")

        # Write everything in the rtf file once the loop is finished
        if len(formatted_rows) > 1:    
            with open(output_rtf_path, 'w', encoding='utf-8') as rtf_file:
                rtf_file.write(rtf_header)
                for row in formatted_rows:
                    rtf_file.write(row)
                rtf_file.write(rtf_footer)
            print(f"Success: RTF file created at {output_rtf_path}")
        else:
            print("Warning: No valid data found to export to RTF")
 
    except Exception as e:
        print(f"Error creating RTF: {e}")

def open_corel():
    try:
        print("attempting to connect to CorelDRAW 2018...")

        # launching the corel application object
        corel = win32com.client.Dispatch("CorelDRAW.Application")
        corel.Visible = True

        # wait for the app to initialize after the bypass
        time.sleep(7)
        print("waiting for UI to stabilize...")

        # attempt to maximize the frame
        try: 
            corel.Frame.WindowState = 1
        except Exception as ui_error:
            print(f"Note: could not set windowstate (UI loading)") 

        # Creating new LTO plate workspace for the print merge
        print("Creating new LTO plate workspace...")
        doc = corel.CreateDocument()

        print("success! corelDraw should now be open")

    except Exception as e:
        print("Failed to open CorelDRAW")
        print(f"Error details: {e}")

def bypass_trial_screen():
    # delay to ensure the evaluation window is the active window
    time.sleep(5)
    print("searching for the trial screen...")

    # sending escape key to close the trial window
    pyautogui.press('esc')
    print('close successfully clicked')

if __name__ == "__main__":
    # execute rtf generation then launch corel
    export_data_to_rtf()
    bypass_trial_screen()
    open_corel()