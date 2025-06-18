def extract_supplier_name(file_name):
    try: 
        # File name related information 
        file_name_split = str(file_name).split("--")
        return file_name_split
    except Exception as e: 
        pass