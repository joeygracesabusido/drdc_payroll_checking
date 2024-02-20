        
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import sys

import math

from prettytable import PrettyTable


from datetime import date


from typing import Optional, List


@staticmethod
def duraville_project():
    
    """This function is for selection of transactions"""
    # print('1001-Search Tons Transaction per Trip Ticket')
    # print('1002-Delete Tonnage Transaction')
   
    # print('x-Exit')

    TransactionList = [
        
            
               {"Code": '2001',"Transaction":'Gross Payroll'},
               {"Code": '2002',"Transaction":'SSS Table'},
               {"Code": '2003',"Transaction":'Payroll computation'},
               {"Code": '2004',"Transaction":'Display Payroll Master List'},
               {"Code": '2005',"Transaction":'Payroll computation 1st Cut-off'},
            
           
           
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

    if ans == '2001':
        return Payrollcomputation.excel_connection()

    elif ans == '2002':
       return Payrollcomputation.excel_connection_sssTable()
    elif ans == '2003':
        return Payrollcomputation.sss_computation()

    elif ans == '2004':
        
        return Payrollcomputation.excel_connection_payroll_masterfile()

    elif ans == '2005':
        return Payrollcomputation.payroll_comp_1st_cut_off()

    elif ans == 'x' or ans =='X':
        exit()

class Payrollcomputation():

    @staticmethod
    def excel_connection_gross_pay():

        sheet_name = 'Payroll-1'

        data_df = pd.read_excel(r'C:\Users\Jerome\Desktop\Payrollcomp\DRDC.xlsx',sheet_name=sheet_name)
       
       

        pd.set_option('display.max_rows', None)

        return(data_df)

        # print(data_df)

    @staticmethod
    def excel_connection_sssTable():

        sheet_name = 'SSS'

        data_df_sss = pd.read_excel(r'C:\Users\Jerome\Desktop\Payrollcomp\DRDC.xlsx',sheet_name=sheet_name)
       
       

        pd.set_option('display.max_rows', None)


        # print(data_df_sss)
        # duraville_project()

        return data_df_sss
    
        

        # print(data_df)
    

    @staticmethod
    def excel_connection_payroll_masterfile(): # this function is for connection of excel file data of Pyaroll master file
        sheet_name = 'PAYROLL-MASTER-FILE'

        data_df_master_file = pd.read_excel(r'C:\Users\Jerome\Desktop\Payrollcomp\DRDC.xlsx',sheet_name=sheet_name)
       
       

        pd.set_option('display.max_rows', None)

        # print(data_df_master_file)

        return data_df_master_file
    
    @staticmethod
    def excel_connection_1st_cut_off(): # this function is for connectio of excel file data of 1st cut-off
        sheet_name = 'PAYROLL-1ST-BATCH'

        data_df_1st_batch = pd.read_excel(r'C:\Users\Jerome\Desktop\Payrollcomp\DRDC.xlsx',sheet_name=sheet_name)
       
       

        pd.set_option('display.max_rows', None)

        # print(data_df_1st_batch)

        return data_df_1st_batch

   
    
    @staticmethod
    def sss_computation():


        payrollData = Payrollcomputation.excel_connection_gross_pay()
        sssData = Payrollcomputation.excel_connection_sssTable()

       

        search = payrollData['Rate']
        # in_range = sssData['Rate_from'],sssData['Rate_to']

        def is_rate_in_range(search):
            return any((search >= sssData['Rate_from']) & (search <= sssData['Rate_to']))
       
        # Apply the function to each rate in search
        in_range = search.apply(is_rate_in_range)


        # Calculate Employee Share based on the condition
        payrollData['Employee Share'] = 0
        payrollData['Employer Share'] = 0
        payrollData['ECC-REMT'] = 0
        payrollData['PHIC'] = 0
        for index, row in sssData.iterrows():
            payrollData.loc[in_range, 'Employee Share'] += (in_range & (payrollData['Rate'] >= row['Rate_from']) & (payrollData['Rate'] <= row['Rate_to'])) * (row['Employee_share'] + row['SS_provident_emp'])
            payrollData.loc[in_range, 'Employer Share'] += (in_range & (payrollData['Rate'] >= row['Rate_from']) & (payrollData['Rate'] <= row['Rate_to'])) * (row['Employer_Share'] + row['SS_provident_empr'])
            payrollData.loc[in_range, 'ECC-REMT'] += (in_range & (payrollData['Rate'] >= row['Rate_from']) & (payrollData['Rate'] <= row['Rate_to'])) * row['ECC'] 
        
         # Handle PHIC calculation
        payrollData['PHIC'] = np.where(search <= 10000, 500 / 2, 0)
        payrollData['PHIC'] += np.where((search > 10000) & (search < 100000.01), search * 0.05 / 2, 0)

        print(payrollData)


        duraville_project()


    @staticmethod
    def payroll_comp_1st_cut_off(): # this function is for computing first cut-off

        master_file = Payrollcomputation.excel_connection_payroll_masterfile()

        cut_off_1st = Payrollcomputation.excel_connection_1st_cut_off()

        semi_monthly_rate = master_file['BASIC_MONTHLY_PAY'] / 2
        employee_id = master_file['EMPLOYEE_ID']

        # Merge the two DataFrames based on the EMPLOYEE_ID column
        merged_data = pd.merge(master_file, cut_off_1st, on='EMPLOYEE_ID', how='inner')

        # Calculate semi-monthly rate
        merged_data['SEMI_MONTHLY_RATE'] = merged_data['BASIC_MONTHLY_PAY'] / 2


        # Calculate GROSS_PAY
        merged_data['GROSS_PAY'] = (merged_data['SEMI_MONTHLY_RATE'] +
                                    merged_data['LATE'] +
                                    merged_data['NORMAL_WORKIG_DAY OT'] +
                                    merged_data['ND_REGULAR_OT'] +
                                    merged_data['TAX_REFUND']).fillna(0)  # fill NaN values with 0
        
        merged_data['TOTAL DEDUCTION'] = (merged_data['SSS_LOAN'] +
                                    merged_data['HDMF_LOAN'].fillna(0) ) # fill NaN values with 0
                                  
        merged_data['NET PAY']  =   merged_data['GROSS_PAY']   -   merged_data['TOTAL DEDUCTION']               


        # Select only the desired columns
        # result_data = merged_data[['EMPLOYEE_ID', 'SEMI_MONTHLY_RATE'] +  list(cut_off_1st.columns)]
        result_data = merged_data[['EMPLOYEE_ID','COMPANY', 'SEMI_MONTHLY_RATE','LATE','NORMAL_WORKIG_DAY OT',
                                   'ND_REGULAR_OT','TAX_REFUND','GROSS_PAY',
                                   'SSS_LOAN', 'HDMF_LOAN','TOTAL DEDUCTION','NET PAY'] ]
        
        


        # Print or return the merged DataFrame
        print(result_data)


        # Print the sum of the GROSS_PAY column
        print("Sum of GROSS_PAY:", result_data['GROSS_PAY'].sum())
        print("Sum of NET_PAY:", result_data['NET PAY'].sum())


        

        
       


        duraville_project()
    
        

        

        

duraville_project()