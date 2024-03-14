import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


from db.mongoDB  import create_mongo_client
mydb = create_mongo_client()


class InsertToMongoDB():
    
    @staticmethod
    def sss_table(): # this is to view the sss table

        sheet_name = 'SSS-TABLE'
        file_path = 'DRDC.xlsx'

        data_df_sss = pd.read_excel(file_path,sheet_name=sheet_name)
       
       

        pd.set_option('display.max_rows', None)


        # print(data_df_sss)
        # print(data_df_sss.dtypes)


        return data_df_sss
    
    @staticmethod
    def insert_sss():

        sss_data = InsertToMongoDB.sss_table()

        print(sss_data)

        ans = input('Do you want to inert SSS table: ').lower()

        if ans == 'yes':

            # Convert DataFrame to dictionary records
            data = sss_data.to_dict(orient='records')
        
            # Insert data into MongoDB
            mydb.sss_table.insert_many(data)

    @staticmethod
    def select_all_sss():

        sss_data = mydb.sss_table.find()

        return sss_data

        # for i in sss_data:
        #     print(i['rate_from'],i['rate_to'])


    @staticmethod
    def sss_data_range():

        sssData = InsertToMongoDB.select_all_sss()

        # Convert list of dictionaries to pandas DataFrame
        # sss_df = pd.DataFrame(sssData)
        # pd.set_option('display.max_rows', None)


        

        sss_data_selection = [{
            "id": i['_id'],
            "rate_from": i['rate_from'],
            "rate_to": i['rate_to'],
            "employee_share": i['employee_share'],
            "ss_provident_emp": i['ss_provident_emp'],
            "employer_Share": i['employer_Share'],
            "ss_provident_empr": i['ss_provident_empr'],
            "ecc": i['ecc'],
        } for i in sssData
        ]
        # print(sss_data_selection)

        salary = input('Enter Salary Range: ')


        for data in sss_data_selection:
            if float(data['rate_from']) <= float(salary) <= data['rate_to']:
                print("Employee share for the given salary range:", data['employee_share'],data['employer_Share'])
                return

        



        # print(sss_df)

        


    
     
InsertToMongoDB.sss_data_range()
       
