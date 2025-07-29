import os

def extract_supplier_name(file_name):
    """
    Extract supplier name from filenames like:
    "Giraffe-Mid Metals Rfp Bidsheet--Giraffe Stainless Cleaned - R1.xlsx"
    
    Returns a tuple: (prefix, supplier_name)
    """
    try:
        # Remove extension and normalize
        name = os.path.splitext(str(file_name))[0].strip()

        # Split by '--' if present
        if "--" in name:
            parts = name.split("--", 1)
            supplier_name = parts[1]
        else:
            supplier_name = name

        # Clean unwanted suffixes
        supplier_name = supplier_name.replace("_r1", "").replace("_r2", "").replace("Cleaned", "").strip()
        supplier_name = supplier_name.replace("_", " ").strip().title()

        return (file_name, supplier_name)
    except Exception as e:
        return (file_name, "Unknown Supplier")
