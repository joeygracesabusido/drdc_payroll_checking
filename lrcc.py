		
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

from os import startfile
import subprocess
import xlsxwriter

from datetime import date


from typing import Optional, List

# from main import main_dashboard


@staticmethod
def lrcc_transaction():
    
    """This function is for selection of transactions"""
    
   
    

    TransactionList = [
        
            
              
               {"Code": '3001',"Transaction":'Payroll computation 1st Cut-off'},
               {"Code": '3002',"Transaction":'Payroll Computation 2nd Cut-off'},
               {"Code": '3003',"Transaction":'Monthly Computation & Govt Mandatory'},
            
           
           
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

    if ans == '3001':
        return Payrollcomputation.payroll_comp_1st_cut_off()

    elif ans == '3002':
        return Payrollcomputation.comp_2nd_cut_off_for_mandatory()
    elif ans == '3003':
        return Payrollcomputation.monthly_computation()
   

    

    elif ans == 'x' or ans =='X':
        exit()

class Payrollcomputation():

    

    @staticmethod
    def excel_connection_sssTable():

        sheet_name = 'SSS'
        file_path = 'LRCC.xlsx'

        data_df_sss = pd.read_excel(file_path,sheet_name=sheet_name)
        # data_df_sss = pd.read_excel(r'/home/joeysabusido/payroll_checking/drdc_payroll_checking/DRDC.xlsx',sheet_name=sheet_name)
       
       

        pd.set_option('display.max_rows', None)


        # print(data_df_sss)
        # duraville_project()

        return data_df_sss
    
        

        # print(data_df)
    

    @staticmethod
    def excel_connection_payroll_masterfile(): # this function is for connection of excel file data of Pyaroll master file
        sheet_name = 'PAYROLL-MASTER-FILE'
        file_path = 'LRCC.xlsx'
        data_df_master_file = pd.read_excel(file_path,sheet_name=sheet_name)
        # data_df_master_file = pd.read_excel(r'/home/joeysabusido/payroll_checking/drdc_payroll_checking/DRDC.xlsx',sheet_name=sheet_name)
       

        pd.set_option('display.max_rows', None)

        # print(data_df_master_file)

        return data_df_master_file
    
    @staticmethod
    def excel_connection_1st_cut_off(): # this function is for connectio of excel file data of 1st cut-off
        sheet_name = 'PAYROLL-1ST-BATCH'
        file_path = 'LRCC.xlsx'
        data_df_1ST_batch = pd.read_excel(file_path,sheet_name=sheet_name)
        # data_df_2ND_batch = pd.read_excel(r'/home/joeysabusido/payroll_checking/drdc_payroll_checking/DRDC.xlsx',sheet_name=sheet_name)
       

        pd.set_option('display.max_rows', None)

        # print(data_df_1st_batch)

        return data_df_1ST_batch
    

    @staticmethod
    def excel_connection_2nd_cut_off(): # this function is for connectio of excel file data of 1st cut-off
        sheet_name = 'PAYROLL-2ND-BATCH'
        file_path = 'LRCC.xlsx'
        data_df_2ND_batch = pd.read_excel(file_path,sheet_name=sheet_name)
        # data_df_1st_batch = pd.read_excel(r'/home/joeysabusido/payroll_checking/drdc_payroll_checking/DRDC.xlsx',sheet_name=sheet_name)
       

        pd.set_option('display.max_rows', None)

        # print(data_df_1st_batch)

        return data_df_2ND_batch

   
    
    @staticmethod
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
        
        


        # Print or return the merged DataFrame
        # print(result_data)


        


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
            workbook = xlsxwriter.Workbook("payroll_comp_1st_cut_off.xlsx")
            worksheet = workbook.add_worksheet('payroll_comp_1st_cut_off')

            # Writing headers
            headers = result_data.columns.tolist()

            for col_num, header in enumerate(headers):
                worksheet.write(0, col_num, header)

            # Writing data
            rowIndex = 1

            for _, row in result_data.iterrows():
                for col_num, value in enumerate(row):
                    worksheet.write(rowIndex, col_num, value)

                rowIndex += 1

            workbook.close()

            # Open the generated Excel file using subprocess
            subprocess.run(['xdg-open', 'payroll_comp_1st_cut_off.xlsx'])


        lrcc_transaction()


    @staticmethod
    def comp_1st_cut_off_for_mandatory():# this function is to return the first cut-off payroll

        master_file = Payrollcomputation.excel_connection_payroll_masterfile()

        cut_off_1st = Payrollcomputation.excel_connection_1st_cut_off()

        
        # Merge the two DataFrames based on the EMPLOYEE_ID column
        merged_data = pd.merge(master_file, cut_off_1st, on='EMPLOYEE_ID', how='inner')

        # Calculate semi-monthly rate
        merged_data['SEMI_MONTHLY_RATE'] = merged_data['BASIC_MONTHLY_PAY'] / 2

         # Replace NaN values with 0 for the entire DataFrame
        merged_data = merged_data.fillna(0)

        
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
                                  
        merged_data['NET PAY']  =   round(merged_data['GROSS_PAY']   -   merged_data['TOTAL DEDUCTION'] , 2)              


        # Select only the desired columns
        # result_data = merged_data[['EMPLOYEE_ID', 'SEMI_MONTHLY_RATE'] +  list(cut_off_1st.columns)]
        result_data = merged_data[['EMPLOYEE_ID','COMPANY','BASIC_MONTHLY_PAY', 'SEMI_MONTHLY_RATE','LATE','NORMAL_WORKIG_DAY OT',
                                   'ND_REGULAR_OT','TAX_REFUND','GROSS_PAY',
                                   'SSS_LOAN', 'HDMF_LOAN','TOTAL DEDUCTION','NET PAY'] ]
        return result_data
    
    @staticmethod
    def comp_2nd_cut_off_for_mandatory():# this function is to return the first cut-off payroll
        # sheet_name = 'PAYROLL-2ND-BATCH'
        # file_path = 'LRCC.xlsx'
        # data_df_2ND_batch = pd.read_excel(file_path,sheet_name=sheet_name)

        master_file = Payrollcomputation.excel_connection_payroll_masterfile()

        cut_off_2nd = Payrollcomputation.excel_connection_2nd_cut_off()
       

       
       
        # Merge the two DataFrames based on the EMPLOYEE_ID column
        merged_data = pd.merge(master_file, cut_off_2nd, on='EMPLOYEE_ID', how='inner')

        # Calculate semi-monthly rate
        merged_data['SEMI_MONTHLY_RATE'] = merged_data['BASIC_MONTHLY_PAY'] / 2

         # Replace NaN values with 0 for the entire DataFrame
        merged_data = merged_data.fillna(0)


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
       
        result_data = merged_data[['EMPLOYEE_ID','COMPANY','BASIC_MONTHLY_PAY', 'SEMI_MONTHLY_RATE','LATE','NORMAL_WORKIG_DAY OT',
                                   'ND_REGULAR_OT','TAX_REFUND','GROSS_PAY'
                                   ] ]
        
        # print(result_data)
        return result_data
    

    @staticmethod
    def monthly_computation(): # this function is for monthly computation

        cut_off_1st = Payrollcomputation.comp_1st_cut_off_for_mandatory()
        cut_off_2nd = Payrollcomputation.comp_2nd_cut_off_for_mandatory()

        # print(cut_off_2nd)  
        sssData = Payrollcomputation.excel_connection_sssTable()
        # print(sssData)

        # Merge the two DataFrames based on the EMPLOYEE_ID column
        merged_data = pd.merge(cut_off_1st, cut_off_2nd, on='EMPLOYEE_ID', how='inner')

        # print(merged_data.columns)

        # Extract desired columns and compute sums
        result_data = merged_data[['EMPLOYEE_ID', 'COMPANY_x',
                                    'SEMI_MONTHLY_RATE_x', 'GROSS_PAY_x', 'GROSS_PAY_y','BASIC_MONTHLY_PAY_y',
                                    'TAX_REFUND_x','TAX_REFUND_y']].groupby(['EMPLOYEE_ID', 'COMPANY_x']).agg({
                                        'SEMI_MONTHLY_RATE_x': 'sum',
                                        'GROSS_PAY_x': 'sum',
                                        'GROSS_PAY_y': 'sum',
                                        'TAX_REFUND_x': 'sum',
                                        'TAX_REFUND_y': 'sum',
                                        'BASIC_MONTHLY_PAY_y': 'sum'
                                        }).reset_index()

        result_data['EMPLOYEE_ID'] = result_data['EMPLOYEE_ID'].astype(str)
        # Calculate total gross pay and select final columns
        result_data['TOTAL_GROSS_PAY'] = round(result_data['GROSS_PAY_x'] + result_data['GROSS_PAY_y'],2)
        # result_data['BASIC_MONTHLY_PAY'] = result_data['BASIC_MONTHLY_PAY']
        
        # Calculate Employee Share based on the condition
        result_data['Employee Share'] = 0
        result_data['Employer Share'] = 0
        result_data['ECC-REMT'] = 0
        result_data['SSS PROVIDENT'] = 0
        result_data['NET TAXABLE'] = 0
        result_data['TAX WITHHELD'] = 0
        result_data['HDMF'] = 0

        # Iterate over each row in sssData
        for index, row_sss in sssData.iterrows():
            # Check if TOTAL_GROSS_PAY is within the range defined by Rate_from and Rate_to
            in_range = (result_data['BASIC_MONTHLY_PAY_y'] >= row_sss['Rate_from']) & (result_data['BASIC_MONTHLY_PAY_y'] <= row_sss['Rate_to'])

            # Calculate Employee Share based on the condition
            result_data.loc[in_range, 'Employee Share'] += in_range * row_sss['Employee_share']
            result_data.loc[in_range, 'Employer Share'] += in_range * row_sss['Employer_Share']
            result_data.loc[in_range, 'SSS PROVIDENT'] += in_range * row_sss['SS_provident_emp']
            result_data.loc[in_range, 'ECC-REMT'] += in_range * row_sss['ECC']

        

        # Handle PHIC calculation
        result_data['PHIC'] = np.where(result_data['BASIC_MONTHLY_PAY_y'] <= 10000, round(500 / 2,2), 0)
        result_data['PHIC'] += np.where((result_data['BASIC_MONTHLY_PAY_y'] > 10000) &
                                        (result_data['BASIC_MONTHLY_PAY_y'] < 100000.01),
                                        round(result_data['BASIC_MONTHLY_PAY_y'] * 0.05 / 2,2), 0)
        result_data['HDMF'] = 100
        result_data['NET TAXABLE'] = round(result_data['TOTAL_GROSS_PAY'] - result_data['Employee Share'] - result_data['PHIC'] - 
                                           result_data['HDMF'] - result_data['SSS PROVIDENT'] - 
                                           result_data['TAX_REFUND_x'] - result_data['TAX_REFUND_y'] , 2)
        
        
        net_taxable_income = result_data['NET TAXABLE']

        def calculate_tax(net_taxable_income):
            # Define your tax brackets and rates
            bracket_1 = (0, 20833)
            bracket_2 = (20833.01, 33332)
            bracket_3 = (33332.01, 66666)
            # ... Add more rates corresponding to the brackets

            # Apply tax rates based on the income brackets
            tax = np.zeros_like(net_taxable_income, dtype=float)

            mask_1 = (net_taxable_income > bracket_1[0]) & (net_taxable_income <= bracket_1[1])
            mask_2 = (net_taxable_income > bracket_2[0]) & (net_taxable_income <= bracket_2[1])
            mask_3 = (net_taxable_income > bracket_3[0]) & (net_taxable_income <= bracket_3[1])

            tax[mask_1] = round(net_taxable_income[mask_1] * 0, 2)
            tax[mask_2] = round((net_taxable_income[mask_2] - 20833) * 0.15, 2)
            tax[mask_3] = round((net_taxable_income[mask_3] - 33332) * 0.20 + 1875, 2)
           

            return tax

        result_data['TAX WITHHELD'] = np.where(result_data['NET TAXABLE'] <= 20833, 0, calculate_tax(result_data['NET TAXABLE']))

        # if result_data['NET TAXABLE'] <= 20800:
        #     result_data['TAX WITHHELD'] = 0

        result_data['NET PAY'] = round(result_data['GROSS_PAY_y'] - result_data['Employee Share'] - result_data['PHIC'] - 
                                           result_data['HDMF'] - result_data['SSS PROVIDENT']- result_data['TAX WITHHELD'], 
                                            2)

        result_data = result_data[['EMPLOYEE_ID', 'TOTAL_GROSS_PAY','GROSS_PAY_y',
                                    'Employee Share','SSS PROVIDENT','Employer Share', 'ECC-REMT', 'PHIC','HDMF','NET TAXABLE','TAX WITHHELD','NET PAY']]

        # Print or return the result_data
        # print(result_data)

        table = PrettyTable(result_data.columns.tolist())  # Convert Index to list
        for _, row in result_data.iterrows():
            table.add_row(row)

        # Print the pretty table
        print(table)
        total_sss = result_data['Employee Share'].sum() + result_data['SSS PROVIDENT'].sum()

        print("Sum of Gross Pay:{:,.2f}".format(result_data['GROSS_PAY_y'].sum()))
        print("Sum of EMPLOYEE SHARES:{:,.2f}".format(total_sss))
        print("Sum of PHIC:{:,.2f}".format(result_data['PHIC'].sum()))
        print("Sum of HDMF:{:,.2f}".format(result_data['HDMF'].sum()))
        print("Sum of NET PAY:{:,.2f}".format(result_data['NET PAY'].sum()))

        ans = input('Do you want to export to excel file ?: ').lower()

        if ans == 'yes':
            workbook = xlsxwriter.Workbook("payroll.xlsx")
            worksheet = workbook.add_worksheet('payroll')

            # Writing headers
            headers = ['EMPLOYEE_ID', 'TOTAL_GROSS_PAY', 'GROSS_PAY', 'Employee Share', 'SSS PROVIDENT',
                       'Employer Share', 'ECC-REMT', 'PHIC', 'HDMF', 'NET TAXABLE', 'TAX WITHHELD', 'NET PAY']

            for col_num, header in enumerate(headers):
                worksheet.write(0, col_num, header)

         

            rowIndex = 2

            for _, row in result_data.iterrows():
                worksheet.write('A' + str(rowIndex), row['EMPLOYEE_ID'])
                worksheet.write('B' + str(rowIndex), row['TOTAL_GROSS_PAY'])
                worksheet.write('C' + str(rowIndex), row['GROSS_PAY_y'])
                worksheet.write('D' + str(rowIndex), row['Employee Share'])
                worksheet.write('E' + str(rowIndex), row['SSS PROVIDENT'])
                worksheet.write('F' + str(rowIndex), row['Employer Share'])
                worksheet.write('G' + str(rowIndex), row['ECC-REMT'])
                worksheet.write('H' + str(rowIndex), row['PHIC'])
                worksheet.write('I' + str(rowIndex), row['HDMF'])
                worksheet.write('J' + str(rowIndex), row['NET TAXABLE'])
                worksheet.write('K' + str(rowIndex), row['TAX WITHHELD'])
                worksheet.write('L' + str(rowIndex), row['NET PAY'])

           
                rowIndex += 1

                # workbook.close()

            workbook.close()

            # Open the generated Excel file using subprocess
            # subprocess.run(['xdg-open', 'payroll.xlsx'])
                # Open the generated Excel file using subprocess
                # subprocess.run(['xdg-open', 'payroll.xlsx'])

            # Open the generated Excel file
            startfile("payroll.xlsx")

        # Create a PDF
            
        # Calculate the sum of NET PAY
            
        total_gross_pay = result_data['GROSS_PAY_y'].sum()
        total_phic = result_data['PHIC'].sum()
        total_hdmf = result_data['HDMF'].sum()
        grand_total_net_pay = result_data['NET PAY'].sum()

        total_monthly_gross = result_data['TOTAL_GROSS_PAY'].sum()
        total_emp_share = result_data['Employee Share'].sum()
        total_sss_provident =  result_data['SSS PROVIDENT'].sum()
       
        total_ecc = result_data['ECC-REMT'].sum()
        total_net_taxable = result_data['NET TAXABLE'].sum()

        total_empr_share = result_data['Employer Share'].sum()

        total_gross_pay = "{:,.2f}".format(total_gross_pay)
        total_phic = "{:,.2f}".format(total_phic)
        total_hdmf = "{:,.2f}".format(total_hdmf)
        grand_total_net_pay = "{:,.2f}".format(grand_total_net_pay)

        total_monthly_gross = "{:,.2f}".format(total_monthly_gross)
        # total_emp_share = "{:,.2f}".format(total_emp_share)
        # total_hdmf = "{:,.2f}".format(total_hdmf)
        # total_sss_provident = "{:,.2f}".format(total_sss_provident)
        # total_ecc = "{:,.2f}".format(total_ecc)
        total_net_taxable = "{:,.2f}".format(total_net_taxable)

        # Append a new row with the grand total
        result_data = result_data._append({'EMPLOYEE_ID': 'Grand Total', 
                                           'NET PAY': grand_total_net_pay,
                                           'GROSS_PAY_y':total_gross_pay,
                                           'PHIC':total_phic,
                                           'HDMF':total_hdmf,
                                           'TOTAL_GROSS_PAY':total_monthly_gross,
                                           'Employee Share':total_emp_share,
                                           'SSS PROVIDENT':total_sss_provident,
                                           'ECC-REMT':total_ecc,
                                           'NET TAXABLE':total_net_taxable,
                                           'Employer Share':total_empr_share}, ignore_index=True)
        pdf_filename = "payroll.pdf"
        
        doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(legal), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)

        # doc = SimpleDocTemplate(pdf_filename, pagesize=letter)

        # Table data
        data = [result_data.columns.tolist()] + result_data.values.tolist()

        # Define the style for the table
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, -1), (-1, -1), colors.red),  # Set text color to red for the last row
        ])

        # Create the table
        table = Table(data, style=style)

        # Build the PDF
        doc.build([table])

        print(f"PDF created successfully: {pdf_filename}")
        # # Open the generated PDF file using the default PDF viewer on Windows
        # subprocess.run(['start', '', pdf_filename], shell=True)
        startfile("payroll.pdf")


        lrcc_transaction()
        
        

        

        

# duraville_project()
