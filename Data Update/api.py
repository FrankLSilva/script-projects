import boto3
import csv
import json
import yaml
import os
import pandas as pd
import copy
import gspread
import config.constants as c
from datetime import datetime, timedelta

class Main:
    def __init__(self):
        try:
            with open(c.PATH_CONFIG, 'r') as f:
                config_data = yaml.safe_load(f)
                
            self.config_data = config_data
        
        except FileNotFoundError:
            print("Arquivo de configuração não encontrado.")
        except yaml.YAMLError as e:
            print(f"Erro ao analisar o arquivo de configuração .yaml: {e}")
        except json.JSONDecodeError as e:
            print(f"Erro ao analisar arquivo JSON de configuração: {e}")
        except Exception as e:
            print(f"Erro ao ler arquivo de configuração: {e}")

class AthenaQueryExecutor(Main):
    def __init__(self):
        try:
            super().__init__()
            self.session = boto3.Session(
                aws_access_key_id= self.config_data.get('aws_access_key_id'),
                aws_secret_access_key= self.config_data.get('aws_secret_access_key'),
                aws_session_token= self.config_data.get('aws_session_token'),
                region_name= self.config_data.get('region_name'))
            self.client = self.session.client('athena')
            self.database = self.config_data.get('database')
            self.output_s3 = self.config_data.get('output_s3')
            self.local_path = self.config_data.get('local_path')
    
        except Exception as e:
            print(f"Erro ao ler arquivo de configuração: {e}")
        
    def execute_query(self, query):
        try:
            response = self.client.start_query_execution(
                QueryString= query,
                QueryExecutionContext={'Database': self.database},
                ResultConfiguration={'OutputLocation': self.output_s3}
            )
            return response['QueryExecutionId']
        except Exception as e:
            print(f"Erro ao executar consulta no Athena: {e}")
    
    def get_query_status(self, query_execution_id):
        response = self.client.get_query_execution(QueryExecutionId= query_execution_id)
        return response['QueryExecution']['Status']['State']

    def paginate_query_results(self, query_execution_id):
        try:
            paginator = self.client.get_paginator('get_query_results')
            for page in paginator.paginate(QueryExecutionId=query_execution_id):
                yield page
        except Exception as e:
            print(f"Erro ao realizar a pginação da consulta: {e}")
       
    def save_results_to_csv(self, query_execution_id, file_name):
        try:
            with open(self.local_path + file_name + '.csv', 'w', encoding=c.UTF_8_ENCODING, newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                for page in self.paginate_query_results(query_execution_id):
                    for row in page['ResultSet']['Rows']:
                        csv_writer.writerow([field.get('VarCharValue', '') for field in row['Data']])
            print(f"Resultados salvos!")
        except Exception as e:
            print(f"Erro ao salvar os resultados da consulta: {e}")
         
class ConnectGoogleSheets(Main):
    def __init__(self):
        super().__init__()
        self.service = c.SERVICE_ACCOUNT
        self.gc = gspread.service_account(filename=self.service)
        self.local_path = self.config_data.get('local_path')
    
    def get_worksheet(self, worksheet):
        return worksheet
    
    def clear_worksheet(self, worksheet):
        worksheet.clear()
        
    def open_spreadsheet(self, spreadsheet):
        return self.gc.open(spreadsheet)
    
    def build_path(self, name_file, extension):
        return os.path.join(self.local_path,name_file+extension)