

import pandas as pd
import numpy as np
import os
import shutil
import sys
import pyfiglet


R = '\033[31m'
G = '\033[32m'
C = '\033[36m'
W = '\033[0m'

splitter = '='*50

def ban():

    print(f'''{pyfiglet.figlet_format('new_cell')}
{splitter}''')

def lister():

    print(f'''
    1.
    2.
    3.
    4.
    5.
    6.
    
{splitter}''')

df_new_CC2 = pd.read_excel('new_cell.xlsx', sheet_name='CC2_Daily')
df_new_CC3 = pd.read_excel('new_cell.xlsx', sheet_name='CC3_Daily')
df_new_RD2 = pd.read_excel('new_cell.xlsx', sheet_name='RD2_Daily')
df_new_RD3 = pd.read_excel('new_cell.xlsx', sheet_name='RD3_Daily')
df_new_RD4 = pd.read_excel('new_cell.xlsx', sheet_name='RD4_Daily')



if __name__ == "__main__":

    while True:

        os.system('cls' if os.name == 'nt' else 'clear')

        ban()

        lister()

        userCh = input('Enter tech as integer : ')

        match userCh:

            case 'CC2':

                pass

            case 'CC3':

                pass

            case 'RD2':

                pass

            case 'RD3':

                pass

            case 'RD4':

                pass




    
