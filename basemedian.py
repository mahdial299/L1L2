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

                    #---------------------------------- 2G voice KPIs
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

                    #---------------------------------- 2G Data KPIs
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

                #==================== data 
                df_CC3 = pd.read_excel('CC3_data.xlsx', sheet_name='Sheet1')
                astro_CC3 = df_CC3.to_dict('records')
                #====================

                main_cell_source_index_3 = df_CC3[['cell_ref']].dropna()
                main_cell_source_index_3 = np.asanyarray(main_cell_source_index_3).flatten()
                main_cell_source_index_3 = list(np.nan_to_num(main_cell_source_index_3))

                for z in range(len(main_cell_source_index_3)):

                    #---------------------------------- 3G Voice KPIs
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

                    for i in range(len(astro_CC3)):

                        if astro_CC3[i]['cell'] == main_cell_source_index_3[z]:

                            kpi_1.append(astro_CC3[i]['cs_erlang'])
                            kpi_2.append(astro_CC3[i]['cs_rrc_connection_establishment_sr'])
                            kpi_3.append(astro_CC3[i]['cs_rab_setup_success_ratio'])
                            kpi_4.append(astro_CC3[i]['softer_handover_success_ratio(hu_cell)'])
                            kpi_5.append(astro_CC3[i]['cs_rab_setup_congestion_rate(hu_cell)'])
                            kpi_6.append(astro_CC3[i]['radio_network_availability_ratio(hu_cell)'])
                            kpi_7.append(astro_CC3[i]['bler_amr(cell_huawei)'])
                            kpi_8.append(astro_CC3[i]['cs_irat_ho_sr'])
                            kpi_9.append(astro_CC3[i]['amr_call_drop_ratio_new(hu_cell)'])
                            kpi_10.append(astro_CC3[i]['csps_rab_setup_success_ratio'])
                            kpi_11.append(astro_CC3[i]['interfrequency_hardhandover_success_ratio_csservice'])
                            kpi_12.append(astro_CC3[i]['cs_cssr'])
                            kpi_13.append(astro_CC3[i]['rrc_setup_success_ratio(cell.service)'])
                            kpi_14.append(astro_CC3[i]['soft_handover_succ_rate'])
                            kpi_15.append(astro_CC3[i]['inter_carrier_ho_success_rate'])
                            kpi_16.append(astro_CC3[i]['cs_rrc_setup_sr_ura_pch(hu_cell)'])
                            kpi_17.append(astro_CC3[i]['cs_cssr_ura_pch(hu_cell)'])

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
       

                    print(f'cell : {C+main_cell_source_index_3[z]+W}')
                    print(Y+'cs_erlang'+W, f'= {kpi_1}'+G,
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'cs_rrc_connection_establishment_sr'+W, f'= {kpi_2}'+G,
                        f'Median : {float(statistics.median(kpi_2))}'+W)
                    print(Y+'cs_rab_setup_success_ratio'+W, f'= {kpi_3}'+G,
                        f'Median : {float(statistics.median(kpi_3))}'+W)
                    print(Y+'softer_handover_success_ratio(hu_cell)'+W, f'= {kpi_4}'+G,
                        f'Median : {float(statistics.median(kpi_4))}'+W)
                    print(Y+'cs_rab_setup_congestion_rate(hu_cell)'+W, f'= {kpi_5}'+G,
                        f'Median : {float(statistics.median(kpi_5))}'+W)
                    print(Y+'radio_network_availability_ratio(hu_cell)'+W, f'= {kpi_6}'+G,
                        f'Median : {float(statistics.median(kpi_6))}'+W)
                    print(Y+'bler_amr(cell_huawei)'+W, f'= {kpi_7}'+G,
                        f'Median : {float(statistics.median(kpi_7))}'+W)
                    print(Y+'cs_irat_ho_sr'+W, f'= {kpi_8}'+G,
                        f'Median : {float(statistics.median(kpi_8))}'+W)
                    print(Y+'amr_call_drop_ratio_new(hu_cell)'+W, f'= {kpi_9}'+G,
                        f'Median : {float(statistics.median(kpi_9))}'+W)
                    print(Y+'csps_rab_setup_success_ratio'+W, f'= {kpi_10}'+G,
                        f'Median : {float(statistics.median(kpi_10))}'+W)
                    print(Y+'interfrequency_hardhandover_success_ratio_csservice'+W, f'= {kpi_11}'+G,
                        f'Median : {float(statistics.median(kpi_11))}'+W)
                    print(Y+'cs_cssr'+W, f'= {kpi_12}'+G,
                        f'Median : {float(statistics.median(kpi_12))}'+W)
                    print(Y+'rrc_setup_success_ratio(cell.service)'+W, f'= {kpi_13}'+G,
                        f'Median : {float(statistics.median(kpi_13))}'+W)
                    print(Y+'soft_handover_succ_rate'+W, f'= {kpi_14}'+G,
                        f'Median : {float(statistics.median(kpi_14))}'+W)
                    print(Y+'inter_carrier_ho_success_rate'+W, f'= {kpi_15}'+G,
                        f'Median : {float(statistics.median(kpi_15))}'+W)
                    print(Y+'cs_rrc_setup_sr_ura_pch(hu_cell)'+W, f'= {kpi_16}'+G,
                        f'Median : {float(statistics.median(kpi_16))}'+W)
                    print(Y+'cs_cssr_ura_pch(hu_cell)'+W, f'= {kpi_17}'+G,
                        f'Median : {float(statistics.median(kpi_17))}'+W)
                    
                    
                    

                print(R+splitter+W)
                userfdec = input(C+'CC3 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()
            
            case 4:
                
                #==================== data 
                df_RD3 = pd.read_excel('RD3_data.xlsx', sheet_name='Sheet1')
                astro_RD3 = df_RD3.to_dict('records')
                #====================

                main_cell_source_index_3 = df_RD3[['cell_ref']].dropna()
                main_cell_source_index_3 = np.asanyarray(main_cell_source_index_3).flatten()
                main_cell_source_index_3 = list(np.nan_to_num(main_cell_source_index_3))

                for z in range(len(main_cell_source_index_3)):

                    #---------------------------------- 2G Data KPIs
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
                    kpi_19 = []
                    kpi_20 = []
                    kpi_21 = []
                    kpi_22 = []
                    kpi_23 = []
                    kpi_24 = []
                    kpi_25 = []
                    kpi_26 = []
                    kpi_27 = []
                    kpi_28 = []
                    kpi_29 = []
                    kpi_30 = []
                    kpi_31 = []
                    kpi_32 = []
                 

                    for i in range(len(astro_RD3)):

                        if astro_RD2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(astro_RD3[i]['payload'])
                            kpi_2.append(astro_RD3[i]['ps_cssr'])
                            kpi_3.append(astro_RD3[i]['ps_call_drop_ratio'])
                            kpi_4.append(astro_RD3[i]['average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'])
                            kpi_5.append(astro_RD3[i]['hsupa_uplink_throughput_in_v16(cell_hu)'])
                            kpi_6.append(astro_RD3[i]['cs+ps_rab_setup_success_ratio'])
                            kpi_7.append(astro_RD3[i]['hsdpa_soft_handover_success_ratio'])
                            kpi_8.append(astro_RD3[i]['hs_share_payload_%'])
                            kpi_9.append(astro_RD3[i]['hsdpa_cdr(%)_(hu_cell)_new'])
                            kpi_10.append(astro_RD3[i]['hsupa_cdr(%)_(hu_cell)_new'])
                            kpi_11.append(astro_RD3[i]['ps_r99_call_drop_ratio_with_pch(hu_cell)'])
                            kpi_12.append(astro_RD3[i]['nack_ratio(cell_huawei)'])
                            kpi_13.append(astro_RD3[i]['hsdpa_scheduling_cell_throughput(cell_huawei)'])
                            kpi_14.append(astro_RD3[i]['hsupa_cell_throughput(kbps)(hu_cell)'])
                            kpi_15.append(astro_RD3[i]['radio_network_availability_ratio(hu_cell)'])
                            kpi_16.append(astro_RD3[i]['ps_rab_setup_success_ratio(hu_cell)'])
                            kpi_17.append(astro_RD3[i]['bler9'])
                            kpi_18.append(astro_RD3[i]['cqi>20'])
                            kpi_19.append(astro_RD3[i]['ps_rrc_connection_success_rate_repeatless(hu_cell)'])
                            kpi_20.append(astro_RD3[i]['ps_r99_rab_setup_success_ratio(hu_cell)'])
                            kpi_21.append(astro_RD3[i]['hsdpa_rab_setup_success_ratio(hu_cell)'])
                            kpi_22.append(astro_RD3[i]['hsupa_rab_setup_success_ratio(hu_cell)'])
                            kpi_23.append(astro_RD3[i]['vs.rab.abnormrel.ps_rnc'])
                            kpi_24.append(astro_RD3[i]['ps_rab_setup_congestion_rate'])
                            kpi_25.append(astro_RD3[i]['ps_rab_setup_success_ratio'])
                            kpi_26.append(astro_RD3[i]['ps_rab_congestion_rate'])
                            kpi_27.append(astro_RD3[i]['hsdpa_user_throughput'])
                            kpi_28.append(astro_RD3[i]['hsupa_throughput_mace'])
                            kpi_29.append(astro_RD3[i]['ps_cssr_ura_pch(hu_cell)'])
                            kpi_30.append(astro_RD3[i]['pch2dch_statetrans_sr(hu_cell)'])
                            kpi_31.append(astro_RD3[i]['mean_rtwp(cell_hu)'])
                            kpi_32.append(astro_RD3[i]['cqi_new(hu_cell)'])


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

                    for item in kpi_19:

                        if str(item) == 'nan':

                            kpi_19.remove(item)

                    for item in kpi_20:

                        if str(item) == 'nan':

                            kpi_20.remove(item)

                    for item in kpi_21:

                        if str(item) == 'nan':

                            kpi_21.remove(item)

                    for item in kpi_22:

                        if str(item) == 'nan':

                            kpi_22.remove(item)

                    for item in kpi_23:

                        if str(item) == 'nan':

                            kpi_23.remove(item)

                    for item in kpi_24:

                        if str(item) == 'nan':

                            kpi_24.remove(item)

                    for item in kpi_25:

                        if str(item) == 'nan':

                            kpi_25.remove(item)

                    for item in kpi_26:

                        if str(item) == 'nan':

                            kpi_26.remove(item)

                    for item in kpi_27:

                        if str(item) == 'nan':

                            kpi_27.remove(item)

                    for item in kpi_28:

                        if str(item) == 'nan':

                            kpi_28.remove(item)

                    for item in kpi_29:

                        if str(item) == 'nan':

                            kpi_29.remove(item)

                    for item in kpi_30:

                        if str(item) == 'nan':

                            kpi_30.remove(item)

                    for item in kpi_31:

                        if str(item) == 'nan':

                            kpi_31.remove(item)

                    for item in kpi_32:

                        if str(item) == 'nan':

                            kpi_32.remove(item)
       

                    print(f'cell : {C+main_cell_source_index_3[z]+W}')
                    print(Y+'payload'+W, f'= {kpi_1}'+G,  
                        f'Median : {float(statistics.median(kpi_1))}'+W)
                    print(Y+'ps_cssr'+W, f'= {kpi_2}'+G,  
                        f'Median : {float(statistics.median(kpi_2))}'+W)
                    print(Y+'ps_call_drop_ratio'+W, f'= {kpi_3}'+G,  
                        f'Median : {float(statistics.median(kpi_3))}'+W)
                    print(Y+'average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'+W, f'= {kpi_4}'+G,  
                        f'Median : {float(statistics.median(kpi_4))}'+W)
                    print(Y+'hsupa_uplink_throughput_in_v16(cell_hu)'+W, f'= {kpi_5}'+G,  
                        f'Median : {float(statistics.median(kpi_5))}'+W)
                    print(Y+'cs+ps_rab_setup_success_ratio'+W, f'= {kpi_6}'+G,  
                        f'Median : {float(statistics.median(kpi_6))}'+W)
                    print(Y+'hsdpa_soft_handover_success_ratio'+W, f'= {kpi_7}'+G,  
                        f'Median : {float(statistics.median(kpi_7))}'+W)
                    print(Y+'hs_share_payload_%'+W, f'= {kpi_8}'+G,  
                        f'Median : {float(statistics.median(kpi_8))}'+W)
                    print(Y+'hsdpa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_9}'+G,  
                        f'Median : {float(statistics.median(kpi_9))}'+W)
                    print(Y+'hsupa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_10}'+G,  
                        f'Median : {float(statistics.median(kpi_10))}'+W)
                    print(Y+'ps_r99_call_drop_ratio_with_pch(hu_cell)'+W, f'= {kpi_11}'+G,  
                        f'Median : {float(statistics.median(kpi_11))}'+W)
                    print(Y+'nack_ratio(cell_huawei)'+W, f'= {kpi_12}'+G,  
                        f'Median : {float(statistics.median(kpi_12))}'+W)
                    print(Y+'hsdpa_scheduling_cell_throughput(cell_huawei)'+W, f'= {kpi_13}'+G,  
                        f'Median : {float(statistics.median(kpi_13))}'+W)
                    print(Y+'hsupa_cell_throughput(kbps)(hu_cell)'+W, f'= {kpi_14}'+G,  
                        f'Median : {float(statistics.median(kpi_14))}'+W)
                    print(Y+'radio_network_availability_ratio(hu_cell)'+W, f'= {kpi_15}'+G,  
                        f'Median : {float(statistics.median(kpi_15))}'+W)
                    print(Y+'ps_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_16}'+G,  
                        f'Median : {float(statistics.median(kpi_16))}'+W)
                    print(Y+'bler9'+W, f'= {kpi_17}'+G,  
                        f'Median : {float(statistics.median(kpi_17))}'+W)
                    print(Y+'cqi>20'+W, f'= {kpi_18}'+G,  
                        f'Median : {float(statistics.median(kpi_18))}'+W)
                    print(Y+'ps_rrc_connection_success_rate_repeatless(hu_cell)'+W, f'= {kpi_19}'+G,  
                        f'Median : {float(statistics.median(kpi_19))}'+W)
                    print(Y+'ps_r99_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_20}'+G,  
                        f'Median : {float(statistics.median(kpi_20))}'+W)
                    print(Y+'hsdpa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_21}'+G,  
                        f'Median : {float(statistics.median(kpi_21))}'+W)
                    print(Y+'hsupa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_22}'+G,  
                        f'Median : {float(statistics.median(kpi_22))}'+W)
                    print(Y+'vs.rab.abnormrel.ps_rnc'+W, f'= {kpi_23}'+G,  
                        f'Median : {float(statistics.median(kpi_23))}'+W)
                    print(Y+'ps_rab_setup_congestion_rate'+W, f'= {kpi_24}'+G,  
                        f'Median : {float(statistics.median(kpi_24))}'+W)
                    print(Y+'ps_rab_setup_success_ratio'+W, f'= {kpi_25}'+G,  
                        f'Median : {float(statistics.median(kpi_25))}'+W)
                    print(Y+'ps_rab_congestion_rate'+W, f'= {kpi_26}'+G,  
                        f'Median : {float(statistics.median(kpi_26))}'+W)
                    print(Y+'hsdpa_user_throughput'+W, f'= {kpi_27}'+G,  
                        f'Median : {float(statistics.median(kpi_27))}'+W)
                    print(Y+'hsupa_throughput_mace'+W, f'= {kpi_28}'+G,  
                        f'Median : {float(statistics.median(kpi_28))}'+W)
                    print(Y+'ps_cssr_ura_pch(hu_cell)'+W, f'= {kpi_29}'+G,  
                        f'Median : {float(statistics.median(kpi_29))}'+W)
                    print(Y+'pch2dch_statetrans_sr(hu_cell)'+W, f'= {kpi_30}'+G,  
                        f'Median : {float(statistics.median(kpi_30))}'+W)
                    print(Y+'mean_rtwp(cell_hu)'+W, f'= {kpi_31}'+G,  
                        f'Median : {float(statistics.median(kpi_31))}'+W)
                    print(Y+'cqi_new(hu_cell)'+W, f'= {kpi_32}'+G,  
                        f'Median : {float(statistics.median(kpi_32))}'+W)

                    
                    
                    

                print(R+splitter+W)
                userfdec = input(C+'RD2 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()

            case 5:
                
                # ==================== data
                df_RD4 = pd.read_excel('RD4_data.xlsx', sheet_name='Sheet1')
                astro_RD4 = df_RD4.to_dict('records')
                #=====================

                main_cell_source_index_2 = df_RD2[['cell_ref']].dropna()
                main_cell_source_index_2 = np.asanyarray(main_cell_source_index_2).flatten()
                main_cell_source_index_2 = list(np.nan_to_num(main_cell_source_index_2))

                for z in range(len(main_cell_source_index_2)):

                    #---------------------------------- 2G Data KPIs
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