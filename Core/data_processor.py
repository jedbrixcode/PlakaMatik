import os

def export_data_to_rtf(input_path, output_path):
    rtf_header = r"{\rtf\ansi\ud\uc1 2\par "
    rtf_footer = r"}"
    formatted_rows = []
    
    headers = r"\\MIDDLE\\\\IDENTIFIER\\\par "
    formatted_rows.append(headers)

    try:
        if not os.path.exists(input_path):
            print(f"crit error: {input_path} not found")
            return False

        with open(input_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            print(f"Reading {len(lines)} lines from input...")

            for line in lines:
                clean_line = line.rstrip()
                
                if not clean_line or "Variable 1" in clean_line:
                    continue

                columns = clean_line.split('    ')

                if len(columns) >= 2:
                    var1 = columns[0].strip()
                    var2 = columns[1].strip()

                    row_string = f"\\\\{var1}\\\\\\\\{var2}\\\\\\par "
                    formatted_rows.append(row_string)
                    print(f"Data parsed: {var1} | {var2}")
                else:
                    print(f"Skipping line (missing 4-space delimiter): {clean_line}")

        if len(formatted_rows) > 1:    
            with open(output_path, 'w', encoding='utf-8') as rtf_file:
                rtf_file.write(rtf_header)
                for row in formatted_rows:
                    rtf_file.write(row)
                rtf_file.write(rtf_footer)
            print(f"Success: RTF file created at {output_path}")
            return True
        else:
            print("Warning: No valid data found to export to RTF")
            return False
 
    except Exception as e:
        print(f"Error creating RTF: {e}")
        return False