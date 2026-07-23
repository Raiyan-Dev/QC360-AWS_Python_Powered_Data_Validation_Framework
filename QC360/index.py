#Script Summary

import os
import sys
import logging
from datetime import datetime
from openpyxl import load_workbook
import pandas as pd
from input_config import get_input_config   
from OVR import OVR
from sales_file import file


# Initialize logging
log_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(log_directory, exist_ok=True)  # Create logs directory if it doesn't exist
log_file_path = os.path.join(log_directory, f"logfile_{datetime.now().strftime('%Y%m%d%H%M%S')}.log")
logging.basicConfig(filename=log_file_path, level=logging.INFO)

try:
    
    # Print the absolute path to the file 'reporting_config.xlsm'
    file_path = os.getcwd()
    file_path = f"{file_path}{os.sep}reporting_config.xlsm"

    print(f"Attempting to access: {file_path}")
    logging.info(f"Attempting to access: {file_path}")

    # Load workbook
    book = load_workbook(file_path, keep_vba=True)

    # Load mapping dataframe from team-mkt-chnl_cfg sheet
    mapping_df = pd.read_excel(file_path, sheet_name='team-mkt-chnl_cfg', skiprows=7, engine='openpyxl')
    mapping_df_intm = mapping_df.iloc[:, 2:]
    

    # Filter rows where 'Process_Active_Flag' is 'Y'
    active_processes_df = mapping_df_intm[mapping_df_intm['Process_Active_Flag'] == 'Y']

    # Extract unique, non-empty, and stripped 'Process' values
    active_processes = active_processes_df['Process'].dropna().unique()
    active_processes = [process.strip() for process in active_processes if process.strip()]

    # Convert active_processes to a list of strings and check if 'ALL' (case-insensitive) is present
    if any(str(process).lower() == 'all' for process in active_processes):
        # If 'ALL' is present, include all possible processes
        active_processes = ['P1_Sales', 'Attainment', 'CLR', 'OVR', 'BS', 'BA', 'HCM', 'Kaiser', 'Pueblo']
    else:
        # Otherwise, keep only the specified processes
        active_processes = list(active_processes)

    # Create mapping list and extract distinct 'Process' values which has Active Flag = Y
    mapping_list = []

    for _, row in mapping_df_intm.iterrows():
        if row['Active Flag'] == 'Y':
            mapping_list.append(row.tolist())

    # Convert list to DataFrame with column names
    mapping_list_df = pd.DataFrame(mapping_list, columns = mapping_df_intm.columns)   

    # Load configuration dataframe
    config_df = pd.read_excel(file_path, sheet_name='Config', skiprows=9, usecols='C:H', nrows=1)
    env = config_df["Environment"][0]
    env_dict = {"Development": 'dev', "Test": 'test', "Production": 'prod'}

    # Get input control configurations df for tables snapshot to be used
    inpt_ctrl_cfg_df = get_input_config(config_df, logging)

    # Display QC Tool header
    print(f"\t\t******************************************************")
    print(f"\t\t*                                                    *")
    print(f"\t\t*   AWS & Python Powered Data Validation Framework   *")
    print(f"\t\t*                                                    *")
    print(f"\t\t******************************************************")
    

    # Extract unique combinations for BU_ID, BU_TAG and QTR_NM from mapping_list_df
    # Filter mapping_list_df to keep only rows where all three columns are not null
    filtered_df = mapping_list_df[mapping_list_df[['BU_ID', 'QTR_NM', 'BU_TAG']].notnull().all(axis=1)]

    # Create unique_bu_qtr_df with all columns 
    unique_bu_qtr_df = filtered_df.drop_duplicates(subset=['BU_ID', 'QTR_NM', 'BU_TAG']).reset_index(drop=True) 

    # Iterate over unique BU_ID and QTR_NM combinations
    for k in range(len(unique_bu_qtr_df)):
    
        #Load queries for each Process sheet
        for process in active_processes:
            
            # Load queries dataframe dynamically based on the Process name (Ex - P1_Sales, Attainment, FSIP_A1 etc)
            queries_df = pd.read_excel(file_path, sheet_name=process, skiprows=3, header=0, engine='openpyxl')
            queries_df.dropna(how= 'all', inplace= True)

            # Intermediate dataframe for creating the filtered dictionary for $$variables
            queries_df_intm = queries_df.iloc[:, 15:]
            filter_dict = {col: f'$${col.lower()}' for col in queries_df_intm.columns}
            
            # Initialize File and Report queries
            queries_df_file = queries_df.loc[
                (queries_df['File/Report'] == 'File') & (queries_df['Active Flag'] == 'Y')
            ].reset_index(drop=True)

            queries_df_OVR_intm = queries_df.loc[
                (queries_df['File/Report'] == 'Report') & (queries_df['Active Flag'] == 'Y')
            ].reset_index(drop=True)

            # Call functions for processing File and Report wise
            OVR(queries_df_OVR_intm, config_df, filter_dict, mapping_list_df, inpt_ctrl_cfg_df, unique_bu_qtr_df.loc[[k]], process, env_dict, env, logging)  #for Report (Excel) generation
            
            file(queries_df_file, config_df, filter_dict, mapping_list_df, inpt_ctrl_cfg_df, unique_bu_qtr_df.loc[[k]], process, env_dict, env, logging)  #for File (.txt, some other delimiter) generation
        
    # Exit program
    sys.exit(0)

except Exception as e:
    print('Following Error Occurred: ')
    print(e)
    logging.error(f'Error: {e}')

    input('Press ENTER to exit')
    sys.exit(0)
