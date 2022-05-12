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

    print(G+f'''1. 2G voice Calculation
2. 2G data Calculation
3. 3G voice Calculation
4. 3G data Calculation
5. 4G data Calculation

6. Excel generation

{R+splitter+W}''')

# ======================================= all data


# df_CC3 = pd.read_excel('CC3_data.xlsx', sheet_name='Sheet1')
# df_RD3 = pd.read_excel('RD3_data.xlsx', sheet_name='Sheet1')
# df_RD4 = pd.read_excel('RD4_data.xlsx', sheet_name='Sheet1')

# astro_CC3 = df_CC3.to_dict('records')
# astro_RD3 = df_RD3.to_dict('records')
# astro_RD4 = df_RD4.to_dict('records')
# ==============================================


if __name__ == '__main__':

    while True:

        os.system('cls' if os.name == 'nt' else 'clear')

        ban()

        lister()

        userCh = int(input('Tech as integer : '))

        match userCh:

            case 1:

                # ================= data
                df_CC2 = pd.read_excel('CC2_data.xlsx', sheet_name='Sheet1')

                astro_CC2 = df_CC2.to_dict('records')
                # ========================

                main_cell_source_index_2 = df_CC2[['cell_ref']].dropna()
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
                    kpi_13 = []
                    kpi_14 = []
                    kpi_15 = []
                    kpi_16 = []
                    kpi_17 = []
                    kpi_18 = []
                   

                    for i in range(len(astro_CC2)):

                        if astro_CC2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(astro_CC2[i]['tch_traffic'])
                            kpi_2.append(astro_CC2[i]['available_tch'])
                            kpi_3.append(astro_CC2[i]['htch_traffic'])
                            kpi_4.append(astro_CC2[i]['sdcch_mht'])
                            kpi_5.append(astro_CC2[i]['tch_availability'])
                            kpi_6.append(astro_CC2[i]['amrfr_usage'])
                            kpi_7.append(astro_CC2[i]['amrhr_usage'])
                            kpi_8.append(astro_CC2[i]['cssr3'])
                            kpi_9.append(astro_CC2[i]['sdcch_congestion_rate'])
                            kpi_10.append(astro_CC2[i]['sdcch_drop_rate'])
                            kpi_11.append(astro_CC2[i]['tch_assignment_fr'])
                            kpi_12.append(astro_CC2[i]['tch_cong'])
                            kpi_13.append(astro_CC2[i]['ihsr2'])
                            kpi_14.append(astro_CC2[i]['ohsr2'])
                            kpi_15.append(astro_CC2[i]['sdcch_access_success_rate2'])
                            kpi_16.append(astro_CC2[i]['cdr3'])
                            kpi_17.append(astro_CC2[i]['rx_qualitty_dl_new'])
                            kpi_18.append(astro_CC2[i]['rx_qualitty_ul_new'])

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

                    for item in kpi_13:

                        if str(item) == 'nan':

                            kpi_13.remove(item)

                    for item in kpi_14:

                        if str(item) == 'nan':

                            kpi_14.remove(item)

                    for item in kpi_15:

                        if str(item) == 'nan':

                            kpi_15.remove(item)

                    for item in kpi_16:

                        if str(item) == 'nan':

                            kpi_16.remove(item)

                    for item in kpi_17:

                        if str(item) == 'nan':

                            kpi_17.remove(item)

                    for item in kpi_18:

                        if str(item) == 'nan':

                            kpi_18.remove(item)
       

                    print(f'cell : {C+main_cell_source_index_2[z]+W}')
                    print(Y+'tch_traffic'+W, f'= {kpi_1}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'available_tch'+W, f'= {kpi_2}'+G,
                        f'Median : {float(statistics.median(kpi_2))}'+W)
                    print(Y+'htch_traffic'+W, f'= {kpi_3}'+G,
                        f'Median : {float(statistics.median(kpi_3))}'+W)
                    print(Y+'sdcch_mht'+W, f'= {kpi_4}'+G,
                        f'Median : {float(statistics.median(kpi_4))}'+W)
                    print(Y+'tch_availability'+W, f'= {kpi_5}'+G,
                        f'Median : {float(statistics.median(kpi_5))}'+W)
                    print(Y+'amrfr_usage'+W, f'= {kpi_6}'+G,
                        f'Median : {float(statistics.median(kpi_6))}'+W)
                    print(Y+'amrhr_usage'+W, f'= {kpi_7}'+G,
                        f'Median : {float(statistics.median(kpi_7))}'+W)
                    print(Y+'cssr3'+W, f'= {kpi_8}'+G,
                        f'Median : {float(statistics.median(kpi_8))}'+W)
                    print(Y+'sdcch_congestion_rate'+W, f'= {kpi_9}'+G,
                        f'Median : {float(statistics.median(kpi_9))}'+W)
                    print(Y+'sdcch_drop_rate'+W, f'= {kpi_10}'+G,
                        f'Median : {float(statistics.median(kpi_10))}'+W)
                    print(Y+'tch_assignment_fr'+W, f'= {kpi_11}'+G,
                        f'Median : {float(statistics.median(kpi_11))}'+W)
                    print(Y+'tch_cong'+W, f'= {kpi_12}'+G,
                        f'Median : {float(statistics.median(kpi_12))}'+W)
                    print(Y+'ihsr2'+W, f'= {kpi_13}'+G,
                        f'Median : {float(statistics.median(kpi_13))}'+W)
                    print(Y+'ohsr2'+W, f'= {kpi_14}'+G,
                        f'Median : {float(statistics.median(kpi_14))}'+W)
                    print(Y+'sdcch_access_success_rate2'+W, f'= {kpi_15}'+G,
                        f'Median : {float(statistics.median(kpi_15))}'+W)
                    print(Y+'cdr3'+W, f'= {kpi_16}'+G,
                        f'Median : {float(statistics.median(kpi_16))}'+W)
                    print(Y+'rx_qualitty_dl_new'+W, f'= {kpi_17}'+G,
                        f'Median : {float(statistics.median(kpi_17))}'+W)
                    print(Y+'rx_qualitty_ul_new'+W, f'= {kpi_18}'+G,
                        f'Median : {float(statistics.median(kpi_18))}'+W)
                    
                    
                    

                print(R+splitter+W)
                userfdec = input(C+'CC2 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()

            case 2:

                # ================= data
                df_RD2 = pd.read_excel('RD2_data.xlsx', sheet_name='Sheet1')

                astro_RD2 = df_RD2.to_dict('records')
                # ========================

                main_cell_source_index_2 = df_RD2[['cell_ref']].dropna()
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

                    for i in range(len(astro_RD2)):

                        if astro_RD2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(astro_RD2[i]['tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'])
                            kpi_2.append(astro_RD2[i]['tbf_drop(ul+dl)(hu_cell)'])
                            kpi_3.append(astro_RD2[i]['average_throughput_of_downlink_gprs_llc_per_user(kbps)'])
                            kpi_4.append(astro_RD2[i]['average_throughput_of_downlink_egprs_llc_per_user(kbps)'])
                            kpi_5.append(astro_RD2[i]['thr_dl_gprs_per_ts(cell_hu)'])
                            kpi_6.append(astro_RD2[i]['thr_dl_egprs_per_ts(cell_hu)'])
                            kpi_7.append(astro_RD2[i]['payload_total_ul(cell_hu)'])
                            kpi_8.append(astro_RD2[i]['payload_total_dl(cell_hu)'])
                            kpi_9.append(astro_RD2[i]['payload_total(cell_hu)'])
                            kpi_10.append(astro_RD2[i]['edge_share_payload(cell_hu)'])
                            kpi_11.append(astro_RD2[i]['tch_availability(hu_cell)'])
                            kpi_12.append(astro_RD2[i]['trx'])

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
                        f'Median : {float(statistics.median(kpi_2))}'+W)
                    print(Y+'average_throughput_of_downlink_gprs_llc_per_user(kbps)'+W, f'= {kpi_3}'+G,
                        f'Median : {float(statistics.median(kpi_3))}'+W)
                    print(Y+'average_throughput_of_downlink_egprs_llc_per_user(kbps)'+W, f'= {kpi_4}'+G,
                        f'Median : {float(statistics.median(kpi_4))}'+W)
                    print(Y+'thr_dl_gprs_per_ts(cell_hu)'+W, f'= {kpi_5}'+G,
                        f'Median : {float(statistics.median(kpi_5))}'+W)
                    print(Y+'thr_dl_egprs_per_ts(cell_hu)'+W, f'= {kpi_6}'+G,
                        f'Median : {float(statistics.median(kpi_6))}'+W)
                    print(Y+'payload_total_ul(cell_hu)'+W, f'= {kpi_7}'+G,
                        f'Median : {float(statistics.median(kpi_7))}'+W)
                    print(Y+'payload_total_dl(cell_hu)'+W, f'= {kpi_8}'+G,
                        f'Median : {float(statistics.median(kpi_8))}'+W)
                    print(Y+'payload_total(cell_hu)'+W, f'= {kpi_9}'+G,
                        f'Median : {float(statistics.median(kpi_9))}'+W)
                    print(Y+'edge_share_payload(cell_hu)'+W, f'= {kpi_10}'+G,
                        f'Median : {float(statistics.median(kpi_10))}'+W)
                    print(Y+'tch_availability(hu_cell)'+W, f'= {kpi_11}'+G,
                        f'Median : {float(statistics.median(kpi_11))}'+W)
                    print(Y+'trx'+W, f'= {kpi_12}'+G,
                        f'Median : {float(statistics.median(kpi_12))}'+W)
                    
                    
                    

                print(R+splitter+W)
                userfdec = input(C+'RD2 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()
        
            case 3:

                pass
            
            case 4:

                pass

            case 5:

                pass