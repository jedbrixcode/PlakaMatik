import os

def parse_input_data(input_path):
    import os
    parsed_data = []

    try:
        if not os.path.exists(input_path):
            print(f"crit error: {input_path} not found")
            return None

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
                    
                    plate_type = None
                    if len(columns) >= 3:
                        plate_type = columns[2].strip()

                    parsed_data.append({
                        "middle": var1,
                        "identifier": var2,
                        "type": plate_type
                    })
                    print(f"Data parsed: {var1} | {var2}")
                else:
                    print(f"Skipping line (missing 4-space delimiter): {clean_line}")

        if len(parsed_data) > 0:    
            print(f"Success: Extracted {len(parsed_data)} valid records.")
            return parsed_data
        else:
            print("Warning: No valid data found.")
            return None
 
    except Exception as e:
        print(f"Error parsing data: {e}")
        return None