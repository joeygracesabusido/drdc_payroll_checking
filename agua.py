        
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import sys
import platform
import math


from prettytable import PrettyTable
from reportlab.lib.pagesizes import letter, landscape,legal
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

from prettytable import PrettyTable

from prettytable import PrettyTable

import subprocess
import xlsxwriter
from os import startfile


from datetime import date

import os


from typing import Optional, List


@staticmethod
def agua_transaction():
    
    """This function is for selection of transactions"""
    
   
    

    TransactionList = [
        
            
              
               {"Code": '4001',"Transaction":'Payroll computation 1st Cut-off'},
               {"Code": '4002',"Transaction":'Monthly Computation & Govt Mandatory'},
            
           
           
        ]
    

    menu = PrettyTable()
    menu.field_names=['Code','Transactions']
        
    
    for x in TransactionList:      
        menu.add_row([
            x['Code'],
            x['Transaction'],
          
        ])
    print(menu)

    ans = input('Please enter code for your Desire transaction: ')

    if ans == '4001':
        return Payrollcomputation.payroll_comp_1st_cut_off()

    elif ans == '4002':
        pass
    #    return Payrollcomputation.excel_connection_sssTable()
    
    elif ans == 'x' or ans =='X':
        exit()


class Payrollcomputation():

    @staticmethod
    def excel_connection_payroll_masterfile(): # this function is for connection of excel file data of Pyaroll master file
       
        folder_path = 'excel_file'
        file_name = 'AGUA.xlsx'
        file_path = os.path.join(folder_path, file_name)
        sheet_name = 'PAYROLL-MASTER-FILE'
        data_df_master_file = pd.read_excel(file_path,sheet_name=sheet_name)
     

        pd.set_option('display.max_rows', None)

        # print(data_df_master_file)

        return data_df_master_file
    
    @staticmethod
    def excel_connection_1st_cut_off(): # this function is for connectio of excel file data of 1st cut-off
        folder_path = 'excel_file'
        file_name = 'AGUA.xlsx'
        file_path = os.path.join(folder_path, file_name)
        
        sheet_name = 'PAYROLL-1ST-BATCH'
        
        data_df_1ST_batch = pd.read_excel(file_path,sheet_name=sheet_name)
        # data_df_2ND_batch = pd.read_excel(r'/home/joeysabusido/payroll_checking/drdc_payroll_checking/DRDC.xlsx',sheet_name=sheet_name)
       

        pd.set_option('display.max_rows', None)

        # print(data_df_1st_batch)

        return data_df_1ST_batch

    def payroll_comp_1st_cut_off(): # this function is for computing first cut-off

        master_file = Payrollcomputation.excel_connection_payroll_masterfile()

        cut_off_1st = Payrollcomputation.excel_connection_1st_cut_off()

        
        # Merge the two DataFrames based on the EMPLOYEE_ID column
        merged_data = pd.merge(master_file, cut_off_1st, on='EMPLOYEE_ID', how='inner')

        # Calculate semi-monthly rate
        merged_data['SEMI_MONTHLY_RATE'] = merged_data['BASIC_MONTHLY_PAY'] / 2

         # Replace NaN values with 0 for the entire DataFrame
        merged_data = merged_data.fillna(0)

        merged_data['EMPLOYEE_ID'] = merged_data['EMPLOYEE_ID'].astype(str)
        # Calculate GROSS_PAY
        merged_data['GROSS_PAY'] = (merged_data['SEMI_MONTHLY_RATE'] +
                                    merged_data['LATE'] +
                                    merged_data['UNDERTIME'] +
                                    merged_data['NORMAL_WORKIG_DAY OT'] +
                                    merged_data['ND_REGULAR_OT'] +
                                    merged_data['SPECIAL_HOLIDAY'] +
                                    merged_data['LEGAL_HOLIDAY'] +
                                    merged_data['ABSENT'] +
                                    merged_data['BASIC_PAY_ADJUSTMENT'] +
                                    merged_data['TAX_REFUND']).fillna(0).round(2)  # fill NaN values with 0
        
        merged_data['TOTAL DEDUCTION'] = (merged_data['SSS_LOAN'] +
                                    merged_data['HDMF_LOAN'].fillna(0) ) # fill NaN values with 0
                                  
        merged_data['NET PAY']  =   round(merged_data['GROSS_PAY']   -   merged_data['TOTAL DEDUCTION'] ,2)              


        # Select only the desired columns
        # result_data = merged_data[['EMPLOYEE_ID', 'SEMI_MONTHLY_RATE'] +  list(cut_off_1st.columns)]
        result_data = merged_data[['EMPLOYEE_ID','SEMI_MONTHLY_RATE','LATE','NORMAL_WORKIG_DAY OT',
                                   'ND_REGULAR_OT','TAX_REFUND','GROSS_PAY',
                                   'SSS_LOAN', 'HDMF_LOAN','TOTAL DEDUCTION','NET PAY'] ]
        
        


      
        # Create a pretty table and add rows
        table = PrettyTable(result_data.columns.tolist())  # Convert Index to list
        for _, row in result_data.iterrows():
            table.add_row(row)

        # Print the pretty table
        print(table)

        # Print the sum of the GROSS_PAY column
        print("Sum of GROSS_PAY:{:,.2f}".format( result_data['GROSS_PAY'].sum()))
        print("Sum of NET_PAY:{:,.2f}".format(result_data['NET PAY'].sum()))
        print("Sum of SSS LOAN:{:,.2f}".format(result_data['SSS_LOAN'].sum()))
        print("Sum of HDMF LOAN:{:,.2f}".format(result_data['HDMF_LOAN'].sum()))


        ans = input('Do you want to export to excel file: ').lower()

        if ans == 'yes':

            result_data.to_excel('agua_payroll.xlsx', index=False)

            # Open the generated Excel file using subprocess

            startfile("agua_payroll.xlsx")
            




        agua_transaction()

# agua_transaction()