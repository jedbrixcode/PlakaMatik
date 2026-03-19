import os

def parse_input_data(input_path):
    import os
    parsed_data = []

    try:
        # check if input file exists
        if not os.path.exists(input_path):
            print(f"crit error: {input_path} not found")
            return None

        # open and read input file
        with open(input_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            print(f"Reading {len(lines)} lines from input...")

            # iterate through each line
            for line in lines:
                clean_line = line.rstrip()
                
                # skip empty lines and header
                if not clean_line or "Variable 1" in clean_line:
                    continue

                # split line into columns
                columns = clean_line.split('    ')

                # check if line has at least 2 columns
                if len(columns) >= 2:
                    var1 = columns[0].strip()
                    var2 = columns[1].strip()
                    
                    # get plate type if available
                    plate_type = None
                    if len(columns) >= 3:
                        plate_type = columns[2].strip()

                    # append parsed data
                    parsed_data.append({
                        "middle": var1,
                        "identifier": var2,
                        "type": plate_type
                    })
                    print(f"Data parsed: {var1} | {var2}")
                else:
                    print(f"Skipping line (missing 4-space delimiter): {clean_line}")

        # check if any data was parsed
        if len(parsed_data) > 0:    
            print(f"Success: Extracted {len(parsed_data)} valid records.")
            return parsed_data
        else:
            print("Warning: No valid data found.")
            return None
 
    except Exception as e:
        print(f"Error parsing data: {e}")
        return None