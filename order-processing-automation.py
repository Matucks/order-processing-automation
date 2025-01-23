import pandas as pd
import os
from datetime import datetime
import warnings
import shutil

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Constantes de Caminho
INPUT_DIR = r"C:\\ProjectAutomation\\Input"
OUTPUT_DIR_BILLED = r"C:\\ProjectAutomation\\Output\\Billed"
OUTPUT_DIR_PROGRESS = r"C:\\ProjectAutomation\\Output\\Progress"
BACKUP_DIR = r"C:\\ProjectAutomation\\Backup"

CURRENT_DATE = datetime.now().strftime("%d-%m-%Y")

def find_first_excel_file(directory):
    for file in os.listdir(directory):
        if file.endswith('.xlsx'):
            return os.path.join(directory, file)
    raise FileNotFoundError("No .xlsx file found in the input directory.")

def extract_branch(value):
    parts = value.split(" - ")
    return parts[1][:3] if len(parts) > 1 else value

def extract_family(value):
    if pd.isna(value):
        return None
    parts = value.split(" - ")
    return parts[1].split()[0] if len(parts) > 1 else None

def filter_billed_orders(dataframe):
    orders_to_exclude = ["ORDER1", "ORDER2"]
    billed_df = dataframe[~dataframe['Order Number'].isin(orders_to_exclude) & 
                          (dataframe['Status'] == 'Billed')].copy()
    billed_df['Event Date'] = pd.to_datetime(billed_df['Event Date'], dayfirst=True, errors='coerce')
    billed_df = billed_df[billed_df['Event Date'] == billed_df['Event Date'].max()]
    billed_df['Branch Code'] = billed_df['Branch Code'].apply(extract_branch)

    columns_to_remove = [
        'Situation', 'Destination Port', 'Region', 'District', 'Dealer Port', 'Interior Color', 
        'Order Type', 'Supply Days', 'Transfer Status', 'Comments', 
        'Process', 'Block Reason', 'Probable Production Week', 'Production Year', 
        'Geographic Division', 'Production Week', 'Plant', 'Customer Code'
    ]
    billed_df.drop(columns=columns_to_remove, errors='ignore', inplace=True)
    for col in ["Destination Unit", "Seller", "Client"]:
        billed_df.insert(billed_df.columns.get_loc("Event Date") + 1, col, "")

    return billed_df

def filter_progress_orders(dataframe):
    orders_to_exclude = ["ORDER1", "ORDER2"]
    progress_df = dataframe[~dataframe['Order Number'].isin(orders_to_exclude) & 
                            (dataframe['Status'] != 'Billed')].copy()
    progress_df['Branch Code'] = progress_df['Branch Code'].apply(extract_branch)
    progress_df['Production Week'] = pd.to_datetime(progress_df['Production Week'], dayfirst=True, errors='coerce')
    progress_df['Probable Production Week'] = pd.to_datetime(progress_df['Probable Production Week'], dayfirst=True, errors='coerce')
    progress_df['Week'] = progress_df[['Production Week', 'Probable Production Week']].max(axis=1)
    progress_df.drop(columns=['Production Week', 'Probable Production Week'], errors='ignore', inplace=True)

    progress_df = progress_df.rename(columns={
        "Model": "MODEL", "Branch Code": "BRANCH", "Order Number": "ORDER", 
        "Exterior Color": "COLOR", "Status": "STATUS", "Options": "OPTIONS"
    })
    progress_df["FAMILY"] = progress_df["MODEL"].apply(extract_family)
    progress_df = progress_df[["FAMILY", "MODEL", "OPTIONS", "COLOR", "BRANCH", "Week", "ORDER", "STATUS"]]
    for col in ["Reservation Date", "Destination Unit", "Seller", "Client"]:
        progress_df.insert(progress_df.columns.get_loc("ORDER") + 1, col, "")

    return progress_df

def save_dataframe_as_table(df, output_path, filename, sheet_name="Sheet1"):
    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, filename)

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        (max_row, max_col) = df.shape
        column_settings = [{'header': column} for column in df.columns]

        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'columns': column_settings,
            'name': sheet_name,
            'style': 'Table Style Medium 9'
        })

    return file_path

try:
    input_file = find_first_excel_file(INPUT_DIR)
    df = pd.read_excel(input_file)

    # Save the Billed Orders file as a table
    df_billed = filter_billed_orders(df)
    billed_file_path = save_dataframe_as_table(df_billed, OUTPUT_DIR_BILLED, f'BILLED_{CURRENT_DATE}.xlsx', sheet_name="Billed")
    print(f"Billed Orders file successfully created as a table: {billed_file_path}")

    # Save the Progress Orders file as a table
    df_progress = filter_progress_orders(df)
    progress_file_path = save_dataframe_as_table(df_progress, OUTPUT_DIR_PROGRESS, f'PROGRESS_{CURRENT_DATE}.xlsx', sheet_name="Progress")
    print(f"Progress Orders file successfully created as a table: {progress_file_path}")

    # Save a copy of the Progress Orders file in the Backup directory
    backup_file_path = save_dataframe_as_table(df_progress, BACKUP_DIR, f'PROGRESS_{CURRENT_DATE}.xlsx', sheet_name="ProgressBackup")
    print(f"Progress Orders file successfully backed up: {backup_file_path}")

except Exception as e:
    print(f"An error occurred: {e}")
