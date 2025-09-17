
ALL_FILES = [
    {
        "consolidation_folder": "bidsheet_other_metal",
        "consolidate_file_name": "bidsheet_other_metal_outlier_consolidate",
        "sheet_name": "3. Bidsheet Other Metals",
    },
    {
        "consolidation_folder": "bidsheet_steel", 
        "consolidate_file_name": "bidsheet_steel_outlier_consolidate", 
        "sheet_name": "2. Bidsheet Steel",
    }, 
    {
        "consolidation_folder": "bidsheet_brass",
        "consolidate_file_name": "bidsheet_brass_outlier_consolidate",  
        "sheet_name": "1. Bidsheet Brass",
    }, 
]
import os
import pandas as pd

for main_files_name_information in ALL_FILES:

    R2_FILES_FOLDER_LOCATION = "./files round 2"
    CLEANED_FILES_FOLDER_LOCATION = f"./cleaned_files/{main_files_name_information['consolidation_folder']}"
    FILES_FOLDER_LOCATION = "./files round 2"

    # files = os.listdir(FILES_FOLDER_LOCATION)
    # for item in files: 
        
        # print("Processing excel sheet ---------------------------------------")
        # print(f"./files/{item}")

        # cleaned_csv_file_name = f"{item.split('.')[0]}_cleaned.xlsx"
        # os.makedirs(CLEANED_FILES_FOLDER_LOCATION, exist_ok=True)
        # df = pd.read_excel(f"./{FILES_FOLDER_LOCATION}/{item}", main_files_name_information["sheet_name"], header=None)
        # header_row_index = None
        # col_start_index = None
        # for index, row in df.iterrows():
        #     for col_index, value in enumerate(row):
        #         if isinstance(value, str) and "ROW ID #" in value:
        #             header_row_index = index
        #             col_start_index = col_index
        #             break
        #     if header_row_index is not None:
        #         break
        # if header_row_index is not None and col_start_index is not None:
        #     headers = df.iloc[header_row_index, col_start_index:].tolist()
            
        #     data_rows = df.iloc[header_row_index + 1:, col_start_index:]
        #     data_rows.columns = headers
        #     data_rows.reset_index(drop=True, inplace=True)
        #     data_rows.to_excel(f"{CLEANED_FILES_FOLDER_LOCATION}/{cleaned_csv_file_name}", index=False)
        #     print(f"Data saved to '{CLEANED_FILES_FOLDER_LOCATION}/{cleaned_csv_file_name}'")
        # else:
            # print("No row containing 'ROW ID #' was found.")

    round2files = os.listdir(R2_FILES_FOLDER_LOCATION)
    for item in round2files:
        print("Processing Round 2 excel sheets ---------------------------------------")
        print(f"./files/{item}")

        cleaned_csv_file_name = f"{item.split('.')[0]}_r2_cleaned.xlsx"
        os.makedirs(CLEANED_FILES_FOLDER_LOCATION, exist_ok=True)
        df = pd.read_excel(f"./files round 2/{item}", main_files_name_information["sheet_name"], header=None)
        header_row_index = None
        col_start_index = None
        for index, row in df.iterrows():
            for col_index, value in enumerate(row):
                if isinstance(value, str) and "ROW ID #" in value:
                    header_row_index = index
                    col_start_index = col_index
                    break
            if header_row_index is not None:
                break
        if header_row_index is not None and col_start_index is not None:
            headers = df.iloc[header_row_index, col_start_index:].tolist()
            
            data_rows = df.iloc[header_row_index + 1:, col_start_index:]
            data_rows.columns = headers
            data_rows.reset_index(drop=True, inplace=True)
            data_rows.to_excel(f"{CLEANED_FILES_FOLDER_LOCATION}/{cleaned_csv_file_name}", index=False)
            print(f"Data saved to '{CLEANED_FILES_FOLDER_LOCATION}/{cleaned_csv_file_name}'")
        else:
            print("No row containing 'ROW ID #' was found.")
