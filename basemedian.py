import pyfiglet
import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import statistics
import os
import sys
from statistics import StatisticsError

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


df_2 = pd.read_excel('RD2_data.xlsx', sheet_name='Sheet1')
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

                main_cell_source_index_2 = df_2[['cell_ref']].dropna()
                main_cell_source_index_2 = np.asanyarray(main_cell_source_index_2).flatten()
                main_cell_source_index_2 = list(np.nan_to_num(main_cell_source_index_2))

                for z in range(len(main_cell_source_index_2)):

                    #---------------------------------- 2G KPIs
                    kpi_1 = []
                    kpi_2 = []
                    kpi_3 = []
                    kpi_4 = []
                    kpi_5 = []
                    kpi_6 = []
                    kpi_7 = []
                    kpi_8 = []
                    kpi_9 = []
                    kpi_10 = []
                    kpi_11 = []
                    kpi_12 = []

                    for i in range(len(astro_2)):

                        if astro_2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(astro_2[i]['tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'])
                            kpi_2.append(astro_2[i]['tbf_drop(ul+dl)(hu_cell)'])
                            kpi_3.append(astro_2[i]['average_throughput_of_downlink_gprs_llc_per_user(kbps)'])
                            kpi_4.append(astro_2[i]['average_throughput_of_downlink_egprs_llc_per_user(kbps)'])
                            kpi_5.append(astro_2[i]['thr_dl_gprs_per_ts(cell_hu)'])
                            kpi_6.append(astro_2[i]['thr_dl_egprs_per_ts(cell_hu)'])
                            kpi_7.append(astro_2[i]['payload_total_ul(cell_hu)'])
                            kpi_8.append(astro_2[i]['payload_total_dl(cell_hu)'])
                            kpi_9.append(astro_2[i]['payload_total(cell_hu)'])
                            kpi_10.append(astro_2[i]['edge_share_payload(cell_hu)'])
                            kpi_11.append(astro_2[i]['tch_availability(hu_cell)'])
                            kpi_12.append(astro_2[i]['trx'])

                        else:

                            continue  

                    for item in kpi_1:

                        if str(item) == 'nan':

                            kpi_1.remove(item)

                    for item in kpi_2:

                        if str(item) == 'nan':

                            kpi_2.remove(item)

                    for item in kpi_3:

                        if str(item) == 'nan':

                            kpi_3.remove(item)

                    for item in kpi_4:

                        if str(item) == 'nan':

                            kpi_4.remove(item)

                    for item in kpi_5:

                        if str(item) == 'nan':

                            kpi_5.remove(item)

                    for item in kpi_6:

                        if str(item) == 'nan':

                            kpi_6.remove(item)

                    for item in kpi_7:

                        if str(item) == 'nan':

                            kpi_7.remove(item)

                    for item in kpi_8:

                        if str(item) == 'nan':

                            kpi_8.remove(item)

                    for item in kpi_9:

                        if str(item) == 'nan':

                            kpi_9.remove(item)

                    for item in kpi_10:

                        if str(item) == 'nan':

                            kpi_10.remove(item)

                    for item in kpi_11:

                        if str(item) == 'nan':

                            kpi_11.remove(item)

                    for item in kpi_12:

                        if str(item) == 'nan':

                            kpi_12.remove(item)
       

                    print(f'cell : {C+main_cell_source_index_2[z]+W}')
                    print(Y+'tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'+W, f'= {kpi_1}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'tbf_drop(ul+dl)(hu_cell)'+W, f'= {kpi_2}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'average_throughput_of_downlink_gprs_llc_per_user(kbps)'+W, f'= {kpi_3}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'average_throughput_of_downlink_egprs_llc_per_user(kbps)'+W, f'= {kpi_4}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'thr_dl_gprs_per_ts(cell_hu)'+W, f'= {kpi_5}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'thr_dl_egprs_per_ts(cell_hu)'+W, f'= {kpi_6}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'payload_total_ul(cell_hu)'+W, f'= {kpi_7}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'payload_total_dl(cell_hu)'+W, f'= {kpi_8}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'payload_total(cell_hu)'+W, f'= {kpi_9}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'edge_share_payload(cell_hu)'+W, f'= {kpi_10}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'tch_availability(hu_cell)'+W, f'= {kpi_11}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'trx'+W, f'= {kpi_12}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    
                    
                    

                print(R+splitter+W)
                userfdec = input(C+'2G calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()
        
