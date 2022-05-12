import pyfiglet
import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import statistics
import os
import sys

R = '\033[31m'
G = '\033[32m'
Y = '\033[33m'
C = '\033[36m'
W = '\033[0m'

splitter = '='*50

def ban():

    print(f'''{C+pyfiglet.figlet_format("L1---L2")}
{R+splitter+W}''')

def lister():

    print(G+f'''2. 2G Calculation
3. 3G Calculation
4. 4G Calculation
{R+splitter+W}''')


df_2 = pd.read_excel('test_data.xlsx', sheet_name='Sheet1')
# df_3 = pd.read_excel('test_data.xlsx', sheet_name='Sheet2')
# df_4 = pd.read_excel('test_data.xlsx', sheet_name='Sheet3')

astro_2 = df_2.to_dict('records')
# astro_3 = df_3.to_dict('records')
# astro_4 = df_4.to_dict('records')


if __name__ == '__main__':

    while True:

        os.system('cls' if os.name == 'nt' else 'clear')

        ban()

        lister()

        userCh = int(input('Tech as integer : '))

        match userCh:

            case 2:

                main_cell_source_index_2 = df_2[['cell_index']].dropna()
                main_cell_source_index_2 = np.asanyarray(main_cell_source_index_2).flatten()
                main_cell_source_index_2 = list(np.nan_to_num(main_cell_source_index_2))

                for z in range(len(main_cell_source_index_2)):

                    #---------------------------------- 2G KPIs
                    kpi_1 = []
                    kpi_2 = []
                    kpi_3 = []

                    for i in range(len(astro_2)):

                        if astro_2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(astro_2[i]['kpi_1'])
                            kpi_2.append(astro_2[i]['kpi_2'])
                            kpi_3.append(astro_2[i]['kpi_3'])

                        else:

                            continue

                    print(f'cell : {C+main_cell_source_index_2[z]+W}')
                    print(Y+'kpi_1'+W, f'= {kpi_1}'+G,
                        f'Median : {int(statistics.median(kpi_1))}'+W)
                    print(Y+'kpi_2'+W, f'= {kpi_2}'+G,
                        f'Median : {int(statistics.median(kpi_2))}'+W)
                    print(Y+'kpi_3'+W, f'= {kpi_3}'+G,
                        f'Median : {int(statistics.median(kpi_3))}'+W)

                print(R+splitter+W)
                userfdec = input(C+'2G calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        os.system('cls' if os.name == 'nt' else 'clear')

                        sys.exit()

            case 3:

                pass

            case 4:

                pass

        
