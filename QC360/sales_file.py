import pandas as pd
from Query import run_query_parallel


def file(queries_df_file, config_df, filter_dict, mapping_list_df, inpt_ctrl_cfg_df, unique_bu_qtr_df, process, env_dict, env, logging):
    for i in range(0, len(queries_df_file)):

        file_path = str(queries_df_file['Path'][i]).strip()
        file_name = str(queries_df_file['File Name'][i]).strip()

        query = queries_df_file["Query"][i].strip()
        for column in queries_df_file:
            query = query.replace("$$env", str(env_dict.get(env)))
            if column == 'Process':
                if queries_df_file[column][i] == 'Goals':
                    query = query.replace(str(filter_dict.get(column)), '_goals_s')
                else:
                    query = query.replace(str(filter_dict.get(column)), '')
            if filter_dict.get(column) is not None and pd.isna(queries_df_file[column])[i] == False:
                query = query.replace(str(filter_dict.get(column)),
                                      str(queries_df_file[column][i]))

        print('Query to be executed = ' + query)
        print('\n')

        df, state = run_query_parallel(query, config_df)

        if state == 'SUCCEEDED':
            if (str(queries_df_file['Delimiter'][i]).strip().lower() == "\\t"):
                df.to_csv(file_path + "\\" + file_name, header=True, index=None,
                          sep='\t')
            else:
                df.to_csv(file_path + "\\" + file_name, header=True, index=None,
                          sep=str(queries_df_file['Delimiter'][i]).strip().lower())

            logging.info("Execution Completed for QC No. = " + str(queries_df_file['QC No.'][i]))

        else:
            print(df)
            print('\n')
            logging.error("Execution failed for QC No. = " + str(queries_df_file['QC No.'][i]))


