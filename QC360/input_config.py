#Script Summary

import io
import boto3
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def get_input_config(config_df, logging):
    # Set up AWS credentials
    access_key = config_df['Access Key'].iloc[0]
    secret_key = config_df['Secret Key'].iloc[0]
    bucket = config_df['Bucket'].iloc[0]
    s3_path = config_df['S3 Path'].iloc[0]
    region = config_df['Region'].iloc[0]
    
    try:
        # Load existing workbook and worksheet
        workbook = load_workbook('reporting_config.xlsm')
        worksheet = workbook['team-mkt-chnl_cfg']

        # Get user-defined data date from cell D5
        user_def_data_dt = worksheet['D5'].value
        user_def_data_dt = user_def_data_dt.strip() if isinstance(user_def_data_dt, str) else user_def_data_dt
        user_def_data_dt = user_def_data_dt if user_def_data_dt else None  #Treat empty strings as None
        
        # Initialize Athena client
        client = boto3.client(
            'athena', 
            region_name=region,
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key
        )
        
        # Construct SQL query for getting the necessary output from inpt_ctrl table
        query = """
        SELECT tbl_nm, attr_col_1, attr_val_1, bu_tag, qtr_nm 
        FROM us_commercial_datalake_fsip_l2_prod.az_jbrm_cfg_fsip_l2_inpt_ctrl_tbl_wkly_obu 
        WHERE data_dt = COALESCE(%s, (SELECT MAX(data_dt) FROM us_commercial_datalake_fsip_l2_prod.az_jbrm_cfg_fsip_l2_inpt_ctrl_tbl_wkly_obu)) 
                AND TRIM(LOWER(CTRL_TYP)) = 'consumption'
                AND TRIM(LOWER(CTRL_PARM)) = 'table-snapshot'
                AND TRIM(LOWER(ATTR_COL_1)) = 'data_dt'
                AND TRIM(LOWER(ACTV_FLG)) = 'y'
                AND TRIM(LOWER(TEAM_NM)) = 'all'
                AND TRIM(LOWER(CTRL_DESC)) IS NULL
        ORDER BY qtr_nm
        """ % (f"'{user_def_data_dt}'" if user_def_data_dt else "NULL")  #Replace user_def_data_dt value
        logging.info("Query ran for fetching data from input control config table: ")
        logging.info(query)

        # Execute the query
        response = client.start_query_execution(
            QueryString=query,
            ResultConfiguration={'OutputLocation': f's3://{bucket}/{s3_path}'}
        )
        query_id = response['QueryExecutionId']
        
        # Poll for query execution status
        execution = client.get_query_execution(QueryExecutionId=query_id)
        status = execution['QueryExecution']['Status']['State']
        
        while status in ['RUNNING', 'QUEUED']:
            execution = client.get_query_execution(QueryExecutionId=query_id)
            status = execution['QueryExecution']['Status']['State']
            if status == 'FAILED':
                raise Exception(execution['QueryExecution']['Status']['StateChangeReason'])
            elif status == 'CANCELLED':
                raise Exception(execution['QueryExecution']['Status']['StateChangeReason'])

        # Read results into DataFrame
        s3_resource = boto3.resource('s3', region_name=region,
                                      aws_access_key_id=access_key,
                                      aws_secret_access_key=secret_key)
        s3_object = s3_resource.Bucket(bucket).Object(key=f'{s3_path}{query_id}.csv').get()
        
        new_data_df = pd.read_csv(io.BytesIO(s3_object['Body'].read()), encoding='utf8')

        # Fill NaN values with a placeholder (like an empty string or a specific value)
        new_data_df['attr_val_1'] = new_data_df['attr_val_1'].fillna('')

        # Convert attr_val_1 to string without decimal
        new_data_df['attr_val_1'] = new_data_df['attr_val_1'].replace('', '0').astype(int).astype(str)

        return new_data_df
    
    except Exception as e:
        print(f"An error occurred: {e}")
        logging.error(f"An error occurred: {e}")
        
    finally:
        workbook.close()