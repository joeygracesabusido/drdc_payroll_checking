import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import sys
import platform
import math

from prettytable import PrettyTable


from datetime import date


from typing import Optional, List

from lpdc import lpdc_transaction
from lrcc import lrcc_transaction

@staticmethod
def main_dashboard(): # this function is for displaying dashboard
    TransactionList = [
        
            
               {"Code": '2000',"Transaction":'LDPC Transactions'},
               {"Code": '3000',"Transaction":'LRCC Transactions'},
             
            
           
           
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

    if ans == '2000':
        return lpdc_transaction()
    
    elif ans == '3000':
        return lrcc_transaction()

    

    elif ans == 'x' or ans =='X':
        return main_dashboard()


main_dashboard()

