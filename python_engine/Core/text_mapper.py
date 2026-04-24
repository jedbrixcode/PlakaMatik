def replace_text_in_shapes(shapes, record, p_type):
    """
    Recursively scans and replaces predefined placeholder labels inside shapes, powerclips, 
    and groups with dynamic dataset records retrieved from the Vue framework loop.
    """
    for i in range(1, shapes.Count + 1):
        s = shapes.Item(i)
        try:
            if hasattr(s, 'Text') and s.Text:
                current_text = s.Text.Story.Text.upper()
                s_name = s.Name.upper() if hasattr(s, 'Name') and s.Name else ""
                
                if p_type.upper() == "MV":
                    if "MIDDLE" in current_text or "<MIDDLE>" in current_text or "MIDDLE" in s_name or "MOCK" in current_text:
                        s.Text.Story.Text = record.get("middle", "")
                    elif "IDENTIFIER" in current_text or "<IDENTIFIER>" in current_text or "IDENTIFIER" in s_name:
                        s.Text.Story.Text = record.get("identifier", "")
                
                elif p_type.upper() == "MC":
                    if "MIDDLE" in current_text or "<MIDDLE>" in current_text or "MIDDLE" in s_name or "MOCK" in current_text:
                        s.Text.Story.Text = record.get("middle", "")
                    elif "IDENTIFIER" in current_text or "<IDENTIFIER>" in current_text or "IDENTIFIER" in s_name:
                        s.Text.Story.Text = record.get("identifier", "")
        except Exception as e:
            pass
        
        try:    
            if s.Type == 7: # cdrGroupShape
                replace_text_in_shapes(s.Shapes, record, p_type)
            elif s.PowerClip:
                replace_text_in_shapes(s.PowerClip.Shapes, record, p_type)
        except:
            pass
