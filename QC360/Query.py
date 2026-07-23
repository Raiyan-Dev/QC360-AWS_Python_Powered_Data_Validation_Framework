#Script Summary

import io
import boto3
import pandas as pd


def run_query_parallel(query, config_df, logging):
    
    access_key = config_df["Access Key"][0]
    secret_key = config_df["Secret Key"][0]
    bucket = config_df["Bucket"][0]
    s3_path = config_df["S3 Path"][0]
    region = config_df["Region"][0]
    query_list=[]
    
    if(type(query) is list):
        query_list=query
    else:
        query_list.append(query)
        
    try:
        client = boto3.client('athena',region_name=region,aws_access_key_id=access_key,aws_secret_access_key=secret_key)
                              
        
        #Starting all the queries in a loop
        for single_query in query_list:
            try:
                queryStart = client.start_query_execution(QueryString=single_query['Query'], ResultConfiguration={'OutputLocation': ('s3://' + bucket + '/' + s3_path)})
                single_query['ExecutionId'] = queryStart['QueryExecutionId']
            except client.exceptions.InvalidRequestException as invalidQuery:
                single_query['ExecutionId'] = None
                single_query['Error'] = str(invalidQuery)
            except client.exceptions.InternalServerException as internServExp:
                raise RuntimeError('Some Error at the Server side') from internServExp
            except client.exceptions.TooManyRequestsException as tooManyRequests:
                raise RuntimeError('Number of requests allowed to the server exhausted') from tooManyRequests


        #Checking the status of all the queries by looping and then recording the result
        for single_query in query_list:
          try:
            #Following block captures the case where There is some syntax error in the query and the query could not be started at all
            if single_query['ExecutionId'] is None:
                single_query['Status'] = 'FAILED'
                single_query['Result'] = None


            else:
                queryStatus = 'QUEUED'
                res = None
                while (queryStatus == 'QUEUED' or queryStatus == 'RUNNING'):
                    response = client.get_query_execution(QueryExecutionId=single_query['ExecutionId'])
                    queryStatus = response['QueryExecution']['Status']['State']
                    if queryStatus == 'FAILED':
                        res = response['QueryExecution']['Status']['StateChangeReason']
                    elif queryStatus == 'CANCELLED':
                        res = response['QueryExecution']['Status']['StateChangeReason']
                    elif queryStatus == 'SUCCEEDED':
                        res = response

                if queryStatus == 'SUCCEEDED':
                    s3_resource = boto3.resource('s3', region_name=region,aws_access_key_id=access_key,aws_secret_access_key=secret_key)
                    s3_object = s3_resource.Bucket(bucket).Object(key=s3_path + single_query['ExecutionId'] + '.csv').get()
                    result = pd.read_csv(io.BytesIO(s3_object['Body'].read()), encoding='utf8')
                    single_query['Status'] = 'SUCCEEDED'
                    single_query['Error'] = None
                    single_query['Result'] = result

                elif queryStatus == 'FAILED':
                    single_query['Status'] = 'FAILED'
                    try:
                        awsResponse=client.batch_get_query_execution(QueryExecutionIds=[str(single_query['ExecutionId']),])
                        queryDetails=awsResponse['QueryExecutions'][0]
                        tempFailResp = str(queryDetails['Status']['StateChangeReason'])
                        single_query['Error'] = tempFailResp
                    except Exception as failedQueryException:
                        single_query['Error'] = "Failed to retrive the error in the query"+str(failedQueryException)
                    single_query['Result'] = None

                else:
                    single_query['Status'] = 'FAILED'
                    single_query['Error'] = 'Some unknown error occured'
                    single_query['Result'] = None


          except client.exceptions.InvalidRequestException as invldReqst:
                single_query['Status'] = 'FAILED'
                single_query['Error'] = 'Invalid Request while fetching the status of the query '+str(invldReqst)
                single_query['Result'] = None

        return query_list
        
    except Exception as e:
        return ('Exception Occured = ' + str(e)), ('Exception Occured = ' + str(e))