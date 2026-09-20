import pandas as pd

def add_unique_job_ids(file_path: str, output_path: str, id_prefix: str = "JOB_"):
    """
    Reads an Excel file and adds a globally unique ID column ('job_id') 
    across all sheets in the workbook.
    """
    excel_file = pd.ExcelFile(file_path)
    sheet_writer = pd.ExcelWriter(output_path, engine='openpyxl')
    
    global_counter = 1  # Global counter across all sheets

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # If job_id already exists, drop it to prevent duplication
        if 'job_id' in df.columns:
            df = df.drop(columns=['job_id'])

        # Generate unique IDs for the current sheet
        sheet_rows = len(df)
        job_ids = [f"{id_prefix}{i:06d}" for i in range(global_counter, global_counter + sheet_rows)]
        
        # Increment counter for subsequent sheets
        global_counter += sheet_rows
        
        # Insert 'job_id' as the first column
        df.insert(0, 'job_id', job_ids)
        
        print(f"Sheet '{sheet_name}': Assigned IDs from {job_ids[0]} to {job_ids[-1]} ({sheet_rows} rows)")
        
        df.to_excel(sheet_writer, sheet_name=sheet_name, index=False)

    sheet_writer.close()
    print(f"\nProcessing complete! File with unique IDs saved as: {output_path}")


# ==========================================
# EXECUTION
# ==========================================
if __name__ == "__main__":
    INPUT_FILE = "merged_postings_with_skills_and_AI.xlsx"
    OUTPUT_FILE = "merged_postings_final.xlsx"
    
    add_unique_job_ids(INPUT_FILE, OUTPUT_FILE)