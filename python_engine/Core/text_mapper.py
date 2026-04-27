def _collect_text_shapes(shapes, result, depth=0):
    """
    Detects text shapes by attempting to access the .Text.Story property directly.
    This is version-agnostic and works regardless of CorelDRAW COM type constants.
    Also recurses into groups (Type == 7).
    """
    for i in range(1, shapes.Count + 1):
        try:
            s = shapes.Item(i)
            t = s.Type
            name = ""
            try:
                name = s.Name
            except:
                pass
            print(f"  {'  ' * depth}Shape[{i}] Type={t} Name='{name}'")

            # Try property-based text detection (version-agnostic)
            try:
                text_content = s.Text.Story.Text
                # If we get here without exception, it IS a text shape
                result.append(s)
                print(f"  {'  ' * depth}  -> Identified as TEXT shape (content: '{text_content[:30]}'...)")
                continue
            except:
                pass

            # Recurse into groups
            if t == 7:
                print(f"  {'  ' * depth}  -> Entering group...")
                _collect_text_shapes(s.Shapes, result, depth + 1)

        except Exception as e:
            print(f"  {'  ' * depth}  Warning: could not inspect shape {i}: {e}")


def replace_text_in_shapes(shapes, record, p_type):
    """
    Injects middle and identifier values into the artistic text objects
    on the Print Layer. Handles grouped shapes recursively.

    MV_PLATE Print Layer: 2 text objects (identifier first, then middle) + 1 rectangle
    MC_PLATE Print Layer: 1 text object (middle) + 1 rectangle

    If values appear swapped, the shapes in your CDR are in a different order —
    just let us know and we'll swap [0] and [1].
    """
    middle_val = record.get("middle", "")
    identifier_val = record.get("identifier", "")

    print(f"  Scanning Print Layer shapes (total on layer: {shapes.Count})...")
    text_shapes = []
    _collect_text_shapes(shapes, text_shapes)
    print(f"  -> Collected {len(text_shapes)} text shape(s) to inject into.")

    try:
        if p_type.upper() == "MV":
            if len(text_shapes) >= 1:
                text_shapes[0].Text.Story.Text = identifier_val
                print(f"  -> Set text_shapes[0] (identifier): '{identifier_val}'")
            if len(text_shapes) >= 2:
                text_shapes[1].Text.Story.Text = middle_val
                print(f"  -> Set text_shapes[1] (middle): '{middle_val}'")
        elif p_type.upper() == "MC":
            if len(text_shapes) >= 1:
                text_shapes[0].Text.Story.Text = middle_val
                print(f"  -> Set text_shapes[0] (middle): '{middle_val}'")
    except Exception as e:
        print(f"  Warning: Text injection error: {e}")

