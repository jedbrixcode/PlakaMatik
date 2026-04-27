import time

def compose_and_align_layers(corel_app, master_doc, master_layers, temp_page, p_type, p_idx, global_dx=0.0, global_dy=0.0):
    """
    Calculates geometrical jigs scaling metrics dynamically and transfers layers optimally
    from template file bindings seamlessly to the A3 export frame structure.
    """
    dx = 0.0
    dy = 0.0
    
    master_layer_payload = master_layers['payload']
    master_layer_mv_guides = master_layers['mv_guides']
    master_layer_bg = master_layers['bg']

    for layer_idx in range(1, temp_page.Layers.Count + 1):
        temp_layer = temp_page.Layers.Item(layer_idx)
        l_name = temp_layer.Name.upper()
        
        sr = corel_app.CreateShapeRange()
        try:
            for i in range(1, temp_layer.Shapes.Count + 1):
                s = temp_layer.Shapes.Item(i)
                if s.Type != 9:
                    sr.Add(s)
        except Exception as le:
            print(f"Warning: Could not fetch layer {l_name}: {le}")
            continue

        if sr.Count > 0:
            try:
                grouped = sr.Group()
                grouped.Copy()
            except:
                # Fallback if only 1 object exists
                sr.Copy()
            
            time.sleep(0.5) 
            master_doc.Activate()
            
            if "PRINT LAYER" in l_name:
                master_layer_payload.Activate()
            elif "GUIDES" in l_name:
                master_layer_mv_guides.Activate()
            else: # "MOCK LAYER" and any other aesthetic layers
                master_layer_bg.Activate()
                
            pasted = master_doc.ActiveLayer.Paste()
            time.sleep(0.5) 

            cx = pasted.PositionX + (pasted.SizeWidth / 2.0)
            cy = pasted.PositionY - (pasted.SizeHeight / 2.0)

            target_px = 42.0 / 2.0 
            
            # Force JIG Consistency Position 1 alignment regardless of Single vs Multiple processing requests
            target_py = 29.7 * 0.75 if p_idx == 0 else 29.7 * 0.25 
            
            dx = target_px - cx + global_dx
            dy = target_py - cy + global_dy
            
            pasted.Move(dx, dy)
