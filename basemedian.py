import pyfiglet
from sqlalchemy import true
import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import sys
import warnings

R = '\033[31m'
G = '\033[32m'
Y = '\033[33m'
C = '\033[36m'
W = '\033[0m'

splitter = '='*50

user = 'Mehdi Alebrahim'

gitdir = fr'C:\Users\{user}\Desktop\L12'


def ban():

    print(f'''{C+pyfiglet.figlet_format("L1---L2")}
{R+splitter+W}''')


def lister():

    print(G+f'''--------------- Quarter calculation ------------
=========================== Daily ==========================

1. 2G voice Calculation
2. 2G data Calculation
3. 3G voice Calculation
4. 3G data Calculation
5. 4G data Calculation

============================ BH ============================

1. 2G voice Calculation
2. 2G data Calculation
3. 3G voice Calculation
4. 3G data Calculation
5. 4G data Calculation


----------------------- Add New Cell

6. 
7.

{R+splitter+W}''')


if __name__ == '__main__':

    while True:

        os.system('cls' if os.name == 'nt' else 'clear')

        ban()

        lister()

        userCh = int(input('Select tech : '))

        match userCh:

            case 1:

                            
        

                outWorkbook = xlsxwriter.Workbook(f"CC2_Daily_BL.xlsx")

                outSheet = outWorkbook.add_worksheet(name='Median_bl')

                # ================= data
                df_CC2 = pd.read_excel('CC2_Daily_data.xlsx', sheet_name='Sheet0')
                astro_CC2 = df_CC2.to_dict('records')
                # ========================

                main_cell_source_index_2 = df_CC2[['cell_ref']].dropna()
                main_cell_source_index_2 = np.asanyarray(
                    main_cell_source_index_2).flatten()
                main_cell_source_index_2 = list(
                    np.nan_to_num(main_cell_source_index_2))

                kpi_list = [

                    'Calculation Period',
                    'Region',
                    'Province',
                    'BSC',
                    'Cell',
                    'CS-Traffic',
                    'CSSR',
                    'AMRHR_Usage',
                    'SDCCH_Cong_Rate',
                    'SDCCH_Drop_Rate',
                    'TCH_Assignment_FR',
                    'IHSR',
                    'OHSR',
                    'SDCCH_Access_Success_Rate',
                    'CDR',
                    'RX_Qualitty_DL',
                    'RX_Qualitty_UL',
                    'Level'

                    

                ]

                for z in range(len(main_cell_source_index_2)):

                    # ---------------------------------- 2G voice KPIs
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


                    for i in range(len(astro_CC2)):



                        if astro_CC2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(astro_CC2[i]['tch_traffic'])
                            kpi_2.append(astro_CC2[i]['cssr3'])
                            kpi_3.append(astro_CC2[i]['amrhr_usage'])
                            kpi_4.append(astro_CC2[i]['sdcch_congestion_rate'])
                            kpi_5.append(astro_CC2[i]['sdcch_drop_rate'])
                            kpi_6.append(astro_CC2[i]['tch_assignment_fr'])
                            kpi_7.append(astro_CC2[i]['ihsr2'])
                            kpi_8.append(astro_CC2[i]['ohsr2'])
                            kpi_9.append(astro_CC2[i]['sdcch_access_success_rate2'])
                            kpi_10.append(astro_CC2[i]['cdr3'])
                            kpi_11.append(astro_CC2[i]['rx_qualitty_dl_new'])
                            kpi_12.append(astro_CC2[i]['rx_qualitty_ul_new'])
                

                        else:

                            continue

                    row_1 = 0

                    row_2 = 0
                    column_2 = 0

                    for kpi in kpi_list:

                        outSheet.write(row_2, column_2, kpi)

                        column_2 += 1


                    try:                   

                        

                        print(f'{z} cell : {C+main_cell_source_index_2[z]+W}')

                        outSheet.write(
                            z + 1, 4, main_cell_source_index_2[z])

                        print(Y+'tch_traffic'+W, f'= {kpi_1}'+G,
                            f'Average : {float(np.nanmean(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 5, float(np.nanmean(kpi_1)))




                        if float(np.nanmean(kpi_1)) > 400:

                            outSheet.write(                     
                            z + 1, 17, 'L1')   

                        elif 250 < float(np.nanmean(kpi_1)) <= 400:

                            outSheet.write(                     
                            z + 1, 17, 'L2')   
                        
                        elif 120 < float(np.nanmean(kpi_1)) <= 250:

                            outSheet.write(                     
                            z + 1, 17, 'L3')   

                        elif 50 < float(np.nanmean(kpi_1)) <= 120:         

                            outSheet.write(                     
                            z + 1, 17, 'L4')   

                        elif float(np.nanmean(kpi_1)) <= 50:

                            outSheet.write(                     
                            z + 1, 17, 'L5')




                        print(Y+'cssr3'+W, f'= {kpi_2}'+G,
                            f'Median : {float(np.nanmedian(kpi_2))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_2))) == True:
                            outSheet.write(
                            z + 1, 6, 'null')        
                        else:
                            outSheet.write(
                            z + 1, 6, float(np.nanmedian(kpi_2)))





                        print(Y+'amrhr_usage'+W, f'= {kpi_3}'+G,
                            f'Median : {float(np.nanmedian(kpi_3))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_3))) == True:
                            outSheet.write(
                            z + 1, 7, 'null')        
                        else:
                            outSheet.write(
                            z + 1, 7, float(np.nanmedian(kpi_3)))





                        print(Y+'sdcch_congestion_rate'+W, f'= {kpi_4}'+G,
                            f'Median : {float(np.nanmedian(kpi_4))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_4))) == True:
                            outSheet.write(
                            z + 1, 8, 'null')        
                        else:
                            outSheet.write(
                            z + 1, 8, float(np.nanmedian(kpi_4)))




                        print(Y+'sdcch_drop_rate'+W, f'= {kpi_5}'+G,
                            f'Median : {float(np.nanmedian(kpi_5))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_5))) == True:
                            outSheet.write(
                            z + 1, 9, 'null')        
                        else:
                            outSheet.write(
                            z + 1, 9, float(np.nanmedian(kpi_5)))





                        print(Y+'tch_assignment_fr'+W, f'= {kpi_6}'+G,
                            f'Median : {float(np.nanmedian(kpi_6))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_6))) == True:
                            outSheet.write(
                            z + 1, 10, 'null')        
                        else:
                            outSheet.write(
                            z + 1, 10, float(np.nanmedian(kpi_6)))



                        print(Y+'ihsr2'+W, f'= {kpi_7}'+G,
                            f'Median : {np.isnan(float(np.nanmedian(kpi_7)))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_7))) == True:
                            outSheet.write(
                            z + 1, 11, 'null')        
                        else:
                            outSheet.write(
                            z + 1, 11, float(np.nanmedian(kpi_7)))



                        print(Y+'ohsr2'+W, f'= {kpi_8}'+G,
                            f'Median : {float(np.nanmedian(kpi_8))}'+W)
                        if np.isnan(float(np.nanmedian(kpi_8))) == True:
                            outSheet.write(
                            z + 1, 12, 'null')   
                        else:
                            outSheet.write(
                            z + 1, 12, float(np.nanmedian(kpi_8)))



                        print(Y+'sdcch_access_success_rate2'+W, f'= {kpi_9}'+G,
                            f'Median : {float(np.nanmedian(kpi_9))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_9))) == True:
                            outSheet.write(
                            z + 1, 13, 'null')   
                        else:
                            outSheet.write(
                            z + 1, 13, float(np.nanmedian(kpi_9)))




                        print(Y+'cdr3'+W, f'= {kpi_10}'+G,
                            f'Median : {float(np.nanmedian(kpi_10))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_10))) == True:
                            outSheet.write(
                            z + 1, 14, 'null')   
                        else:
                            outSheet.write(
                            z + 1, 14, float(np.nanmedian(kpi_10)))




                        print(Y+'rx_qualitty_dl_new'+W, f'= {kpi_11}'+G,
                            f'Median : {float(np.nanmedian(kpi_11))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_11))) == True:
                            outSheet.write(
                            z + 1, 15, 'null')   
                        else:
                            outSheet.write(
                            z + 1, 15, float(np.nanmedian(kpi_11)))




                        print(Y+'rx_qualitty_ul_new'+W, f'= {kpi_12}'+G,
                            f'Median : {float(np.nanmedian(kpi_12))}'+W)

                        if np.isnan(float(np.nanmedian(kpi_12))) == True:
                            outSheet.write(
                            z + 1, 16, 'null')   
                        else:
                            outSheet.write(
                            z + 1, 16, float(np.nanmedian(kpi_12)))
                        
            
                        

                    except(TypeError):

                        continue

                print(R+splitter+W)

                outWorkbook.close()

                userfdec = input(
                    C+'CC2 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()

            case 2:

                # os.makedirs(fr'{gitdir}/RD2_BL')
                outWorkbook = xlsxwriter.Workbook(f"RD2_Daily_BL.xlsx")

                outSheet = outWorkbook.add_worksheet(name='Median_bl')

                # ================= data
                df_RD2 = pd.read_excel('RD2_Daily_data.xlsx', sheet_name='Sheet0')
                astro_RD2 = df_RD2.to_dict('records')
                # ========================

                main_cell_source_index_2 = df_RD2[['cell_ref']].dropna()
                main_cell_source_index_2 = np.asanyarray(
                    main_cell_source_index_2).flatten()
                main_cell_source_index_2 = list(
                    np.nan_to_num(main_cell_source_index_2))

                kpi_list = [

                    'payload_total(cell_hu)',
                    'Level',
                    'tbf_establishment_success_rate(ul+dl)(%)(hu_cell)',
                    'tbf_drop(ul+dl)(hu_cell)',
                    'average_throughput_of_downlink_gprs_llc_per_user(kbps)',
                    'average_throughput_of_downlink_egprs_llc_per_user(kbps)',
                    'thr_dl_gprs_per_ts(cell_hu)',
                    'thr_dl_egprs_per_ts(cell_hu)',
                    'edge_share_payload(cell_hu)'

                ]

                for z in range(len(main_cell_source_index_2)):

                    # ---------------------------------- 2G Data KPIs
                    kpi_1 = []
                    kpi_2 = []
                    kpi_3 = []
                    kpi_4 = []
                    kpi_5 = []
                    kpi_6 = []
                    kpi_7 = []
                    kpi_8 = []
               

                    for i in range(len(astro_RD2)):

                        if astro_RD2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(
                                astro_RD2[i]['payload_total(cell_hu)'])
                            kpi_2.append(
                                astro_RD2[i]['tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'])
                            kpi_3.append(
                                astro_RD2[i]['tbf_drop(ul+dl)(hu_cell)'])
                            kpi_4.append(
                                astro_RD2[i]['average_throughput_of_downlink_gprs_llc_per_user(kbps)'])
                            kpi_5.append(
                                astro_RD2[i]['average_throughput_of_downlink_egprs_llc_per_user(kbps)'])
                            kpi_6.append(
                                astro_RD2[i]['thr_dl_gprs_per_ts(cell_hu)'])
                            kpi_7.append(
                                astro_RD2[i]['thr_dl_egprs_per_ts(cell_hu)'])
                            kpi_8.append(
                                astro_RD2[i]['edge_share_payload(cell_hu)'])
                         

                        else:

                            continue

                    # ================================== excel writing main
                    row_1 = 0

                    row_2 = 0
                    column_2 = 1

                    for kpi in kpi_list:

                        outSheet.write(row_2, column_2, kpi)

                        column_2 += 1

                    outSheet.write(row_2, row_1, 'cell')

                    try:

                        

                        print(f'{z} cell : {C+main_cell_source_index_2[z]+W}')
                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_2[z])
                        
                        

                        print(Y+'payload_total(cell_hu)'+W, f'= {kpi_1}'+G,
                              f'Average : {float(np.nanmean(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 1, float(np.nanmean(kpi_1)))

                        if float(np.nanmean(kpi_1)) > 1:

                            outSheet.write(                     
                            z + 1, 2, 'L1')   

                        elif 0.7 < float(np.nanmean(kpi_1)) <= 1:

                            outSheet.write(                     
                            z + 1, 2, 'L2')   
                        
                        elif 0.4 < float(np.nanmean(kpi_1)) <= 0.7:

                            outSheet.write(                     
                            z + 1, 2, 'L3')   

                        elif 0.2 < float(np.nanmean(kpi_1)) <= 0.4:         

                            outSheet.write(                     
                            z + 1, 2, 'L4')   

                        elif float(np.nanmean(kpi_1)) <= 0.2:

                            outSheet.write(                     
                            z + 1, 2, 'L5')

                        print(Y+'tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'+W, f'= {kpi_2}'+G,
                              f'Median : {float(np.nanmedian(kpi_2))}'+W)

                        outSheet.write(
                            z + 1, 3, float(np.nanmedian(kpi_2)))

                        print(Y+'tbf_drop(ul+dl)(hu_cell)'+W, f'= {kpi_3}'+G,
                              f'Median : {float(np.nanmedian(kpi_3))}'+W)

                        outSheet.write(
                            z + 1, 4, float(np.nanmedian(kpi_3)))

                        print(Y+'average_throughput_of_downlink_gprs_llc_per_user(kbps)'+W, f'= {kpi_4}'+G,
                              f'Median : {float(np.nanmedian(kpi_4))}'+W)

                        outSheet.write(
                            z + 1, 5, float(np.nanmedian(kpi_4)))

                        print(Y+'average_throughput_of_downlink_egprs_llc_per_user(kbps)'+W, f'= {kpi_5}'+G,
                              f'Median : {float(np.nanmedian(kpi_5))}'+W)

                        outSheet.write(
                            z + 1, 6, float(np.nanmedian(kpi_5)))

                        print(Y+'thr_dl_gprs_per_ts(cell_hu)'+W, f'= {kpi_6}'+G,
                              f'Median : {float(np.nanmedian(kpi_6))}'+W)

                        outSheet.write(
                            z + 1, 7, float(np.nanmedian(kpi_6)))

                        print(Y+'thr_dl_egprs_per_ts(cell_hu)'+W, f'= {kpi_7}'+G,
                              f'Median : {float(np.nanmedian(kpi_7))}'+W)

                        outSheet.write(
                            z + 1, 8, float(np.nanmedian(kpi_7)))

                        print(Y+'edge_share_payload(cell_hu)'+W, f'= {kpi_8}'+G,
                              f'Median : {float(np.nanmedian(kpi_8))}'+W)

                        outSheet.write(
                            z + 1, 9, float(np.nanmedian(kpi_8)))

                            # warnings.filterwarnings(action='ignore', message='All-NaN slice encountered')

        

                    except(TypeError):

                        continue

                print(R+splitter+W)

                outWorkbook.close()

                userfdec = input(
                    C+'RD2 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()

            case 3:

                outWorkbook = xlsxwriter.Workbook(f"CC3_Daily_BL.xlsx")

                outSheet = outWorkbook.add_worksheet(name='Median_bl')

                # ==================== data
                df_CC3 = pd.read_excel('CC3_Daily_data.xlsx', sheet_name='Sheet0')
                astro_CC3 = df_CC3.to_dict('records')
                # ====================

                main_cell_source_index_3 = df_CC3[['cell_ref']].dropna()
                main_cell_source_index_3 = np.asanyarray(
                    main_cell_source_index_3).flatten()
                main_cell_source_index_3 = list(
                    np.nan_to_num(main_cell_source_index_3))

                kpi_list = [

                    'cs_erlang',
                    'Level',
                    'cs_rab_setup_success_ratio',
                    'softer_handover_success_ratio(hu_cell)',
                    'cs_irat_ho_sr',
                    'amr_call_drop_ratio_new(hu_cell)',
                    'interfrequency_hardhandover_success_ratio_csservice',
                    'cs_rrc_connection_establishment_sr'
                    

                ]
                for z in range(len(main_cell_source_index_3)):

                    # ---------------------------------- 3G Voice KPIs
                    kpi_1 = []
                    kpi_2 = []
                    kpi_3 = []
                    kpi_4 = []
                    kpi_5 = []
                    kpi_6 = []
                    kpi_7 = []
                  
             

                    for i in range(len(astro_CC3)):

                        if astro_CC3[i]['cell'] == main_cell_source_index_3[z]:

                            kpi_1.append(astro_CC3[i]['cs_erlang'])
                            kpi_2.append(
                                astro_CC3[i]['cs_rab_setup_success_ratio'])
                            kpi_3.append(
                                astro_CC3[i]['softer_handover_success_ratio(hu_cell)'])
                            kpi_4.append(
                                astro_CC3[i]['cs_irat_ho_sr'])
                            kpi_5.append(
                                astro_CC3[i]['amr_call_drop_ratio_new(hu_cell)'])
                            kpi_6.append(
                                astro_CC3[i]['interfrequency_hardhandover_success_ratio_csservice'])
                            kpi_7.append(
                                astro_CC3[i]['cs_rrc_connection_establishment_sr'])
                            

                        else:

                            continue

                    row_1 = 0

                    row_2 = 0
                    column_2 = 1

                    for kpi in kpi_list:

                        outSheet.write(row_2, column_2, kpi)

                        column_2 += 1

                    outSheet.write(row_2, row_1, 'cell')

                    try:

                        print(f'{z} cell : {C+main_cell_source_index_3[z]+W}')
                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_3[z])

                        print(Y+'cs_erlang'+W, f'= {kpi_1}'+G,
                              f'Average : {float(np.nanmean(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 1, float(np.nanmean(kpi_1)))

                        if float(np.nanmean(kpi_1)) > 150:

                            outSheet.write(                     
                            z + 1, 2, 'L1')   

                        elif 120 < float(np.nanmean(kpi_1)) <= 150:

                            outSheet.write(                     
                            z + 1, 2, 'L2')   
                        
                        elif 90 < float(np.nanmean(kpi_1)) <= 120:

                            outSheet.write(                     
                            z + 1, 2, 'L3')   

                        elif 50 < float(np.nanmean(kpi_1)) <= 90:         

                            outSheet.write(                     
                            z + 1, 2, 'L4')   

                        elif float(np.nanmean(kpi_1)) <= 50:

                            outSheet.write(                     
                            z + 1, 2, 'L5')

                        print(Y+'cs_rab_setup_success_ratio'+W, f'= {kpi_2}'+G,
                              f'Median : {float(np.nanmedian(kpi_2))}'+W)

                        outSheet.write(
                            z + 1, 3, float(np.nanmedian(kpi_2)))

                        print(Y+'softer_handover_success_ratio(hu_cell)'+W, f'= {kpi_3}'+G,
                              f'Median : {float(np.nanmedian(kpi_3))}'+W)

                        outSheet.write(
                            z + 1, 4, float(np.nanmedian(kpi_3)))

                        print(Y+'cs_irat_ho_sr)'+W, f'= {kpi_4}'+G,
                              f'Median : {float(np.nanmedian(kpi_4))}'+W)

                        outSheet.write(
                            z + 1, 5, float(np.nanmedian(kpi_4)))

                        print(Y+'amr_call_drop_ratio_new(hu_cell)'+W, f'= {kpi_5}'+G,
                              f'Median : {float(np.nanmedian(kpi_5))}'+W)

                        outSheet.write(
                            z + 1, 6, float(np.nanmedian(kpi_5)))

                        print(Y+'interfrequency_hardhandover_success_ratio_csservice'+W, f'= {kpi_6}'+G,
                              f'Median : {float(np.nanmedian(kpi_6))}'+W)

                        outSheet.write(
                            z + 1, 7, float(np.nanmedian(kpi_6)))

                        print(Y+'cs_rrc_connection_establishment_sr'+W, f'= {kpi_7}'+G,
                              f'Median : {float(np.nanmedian(kpi_7))}'+W)

                        outSheet.write(
                            z + 1, 8, float(np.nanmedian(kpi_7)))
                        



                    except(TypeError):

                        continue

                    

                print(R+splitter+W)

                outWorkbook.close()

                userfdec = input(
                    C+'CC3 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()

            case 4:

                outWorkbook = xlsxwriter.Workbook(f"RD3_Daily_BL.xlsx")

                outSheet = outWorkbook.add_worksheet(name='Median_bl')

                # ==================== data
                df_RD3 = pd.read_excel('RD3_Daily_data.xlsx', sheet_name='Sheet0')
                astro_RD3 = df_RD3.to_dict('records')
                # ====================

                main_cell_source_index_3 = df_RD3[['cell_ref']].dropna()
                main_cell_source_index_3 = np.asanyarray(
                    main_cell_source_index_3).flatten()
                main_cell_source_index_3 = list(
                    np.nan_to_num(main_cell_source_index_3))

                kpi_list = [
                    
                    'payload',
                    'Level',
                    'ps_cssr',
                    'ps_call_drop_ratio',
                    'average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)',
                    'hsupa_uplink_throughput_in_v16(cell_hu)',
                    'cs+ps_rab_setup_success_ratio',
                    'hsdpa_soft_handover_success_ratio',
                    'hs_share_payload_%',
                    'hsdpa_cdr(%)_(hu_cell)_new',
                    'hsdpa_scheduling_cell_throughput(cell_huawei)',
                    'ps_rrc_connection_success_rate_repeatless(hu_cell)',
                    'hsdpa_rab_setup_success_ratio(hu_cell)',
                    'hsupa_rab_setup_success_ratio(hu_cell)',
                    'ps_rab_setup_success_ratio',
                    'hsupa_cdr(%)_(hu_cell)_new'



                ]

                for z in range(len(main_cell_source_index_3)):

                    # ---------------------------------- 2G Data KPIs
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
          
                    for i in range(len(astro_RD3)):

                        if astro_RD3[i]['cell'] == main_cell_source_index_3[z]:

                            kpi_1.append(astro_RD3[i]['payload'])
                            kpi_2.append(astro_RD3[i]['ps_cssr'])
                            kpi_3.append(astro_RD3[i]['ps_call_drop_ratio'])
                            kpi_4.append(
                                astro_RD3[i]['average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'])
                            kpi_5.append(
                                astro_RD3[i]['hsupa_uplink_throughput_in_v16(cell_hu)'])
                            kpi_6.append(
                                astro_RD3[i]['cs+ps_rab_setup_success_ratio'])
                            kpi_7.append(
                                astro_RD3[i]['hsdpa_soft_handover_success_ratio'])
                            kpi_8.append(astro_RD3[i]['hs_share_payload_%'])
                            kpi_9.append(
                                astro_RD3[i]['hsdpa_cdr(%)_(hu_cell)_new'])
                            kpi_10.append(
                                astro_RD3[i]['hsdpa_scheduling_cell_throughput(cell_huawei)'])
                            kpi_11.append(
                                astro_RD3[i]['ps_rrc_connection_success_rate_repeatless(hu_cell)'])
                            kpi_12.append(
                                astro_RD3[i]['hsdpa_rab_setup_success_ratio(hu_cell)'])
                            kpi_13.append(
                                astro_RD3[i]['hsupa_rab_setup_success_ratio(hu_cell)'])
                            kpi_14.append(
                                astro_RD3[i]['ps_rab_setup_success_ratio'])
                            kpi_15.append(
                                astro_RD3[i]['hsupa_cdr(%)_(hu_cell)_new'])
                         

                        else:

                            continue

                    row_1 = 0

                    row_2 = 0
                    column_2 = 1

                    for kpi in kpi_list:

                        outSheet.write(row_2, column_2, kpi)

                        column_2 += 1

                    outSheet.write(row_2, row_1, 'cell')

                    try:

                        print(f'{z} cell : {C+main_cell_source_index_3[z]+W}')
                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_3[z])

                        print(Y+'payload'+W, f'= {kpi_1}'+G,
                              f'Average : {float(np.nanmean(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 1, float(np.nanmean(kpi_1)))

                        
                        if float(np.nanmean(kpi_1)) > 28:

                            outSheet.write(                     
                            z + 1, 2, 'L1')   

                        elif 23 < float(np.nanmean(kpi_1)) <= 28:

                            outSheet.write(                     
                            z + 1, 2, 'L2')   
                        
                        elif 18 < float(np.nanmean(kpi_1)) <= 23:

                            outSheet.write(                     
                            z + 1, 2, 'L3')   

                        elif 13 < float(np.nanmean(kpi_1)) <= 18:         

                            outSheet.write(                     
                            z + 1, 2, 'L4')   

                        elif float(np.nanmean(kpi_1)) <= 13:

                            outSheet.write(                     
                            z + 1, 2, 'L5')

                        print(Y+'ps_cssr'+W, f'= {kpi_2}'+G,
                              f'Median : {float(np.nanmedian(kpi_2))}'+W)

                        outSheet.write(
                            z + 1, 3, float(np.nanmedian(kpi_2)))

                        print(Y+'ps_call_drop_ratio'+W, f'= {kpi_3}'+G,
                              f'Median : {float(np.nanmedian(kpi_3))}'+W)

                        outSheet.write(
                            z + 1, 4, float(np.nanmedian(kpi_3)))

                        print(Y+'average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'+W, f'= {kpi_4}'+G,
                              f'Median : {float(np.nanmedian(kpi_4))}'+W)

                        outSheet.write(
                            z + 1, 5, float(np.nanmedian(kpi_4)))

                        print(Y+'hsupa_uplink_throughput_in_v16(cell_hu)'+W, f'= {kpi_5}'+G,
                              f'Median : {float(np.nanmedian(kpi_5))}'+W)

                        outSheet.write(
                            z + 1, 6, float(np.nanmedian(kpi_5)))

                        print(Y+'cs+ps_rab_setup_success_ratio'+W, f'= {kpi_6}'+G,
                              f'Median : {float(np.nanmedian(kpi_6))}'+W)

                        outSheet.write(
                            z + 1, 7, float(np.nanmedian(kpi_6)))

                        print(Y+'hsdpa_soft_handover_success_ratio'+W, f'= {kpi_7}'+G,
                              f'Median : {float(np.nanmedian(kpi_7))}'+W)

                        outSheet.write(
                            z + 1, 8, float(np.nanmedian(kpi_7)))

                        print(Y+'hs_share_payload_%'+W, f'= {kpi_8}'+G,
                              f'Median : {float(np.nanmedian(kpi_8))}'+W)

                        outSheet.write(
                            z + 1, 9, float(np.nanmedian(kpi_8)))

                        print(Y+'hsdpa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_9}'+G,
                              f'Median : {float(np.nanmedian(kpi_9))}'+W)

                        outSheet.write(
                            z + 1, 10, float(np.nanmedian(kpi_9)))

                        print(Y+'hsdpa_scheduling_cell_throughput(cell_huawei)'+W, f'= {kpi_10}'+G,
                              f'Median : {float(np.nanmedian(kpi_10))}'+W)

                        outSheet.write(
                            z + 1, 11, float(np.nanmedian(kpi_10)))

                        print(Y+'ps_rrc_connection_success_rate_repeatless(hu_cell)'+W, f'= {kpi_11}'+G,
                              f'Median : {float(np.nanmedian(kpi_11))}'+W)

                        outSheet.write(
                            z + 1, 12, float(np.nanmedian(kpi_11)))

                        print(Y+'hsdpa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_12}'+G,
                              f'Median : {float(np.nanmedian(kpi_12))}'+W)

                        outSheet.write(
                            z + 1, 13, float(np.nanmedian(kpi_12)))

                        print(Y+'hsupa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_13}'+G,
                              f'Median : {float(np.nanmedian(kpi_13))}'+W)

                        outSheet.write(
                            z + 1, 14, float(np.nanmedian(kpi_13)))

                        print(Y+'ps_rab_setup_success_ratio'+W, f'= {kpi_14}'+G,
                              f'Median : {float(np.nanmedian(kpi_14))}'+W)

                        outSheet.write(
                            z + 1, 15, float(np.nanmedian(kpi_14)))

                        print(Y+'hsupa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_15}'+G,
                              f'Median : {float(np.nanmedian(kpi_15))}'+W)

                        outSheet.write(
                            z + 1, 16, float(np.nanmedian(kpi_15)))

                      


                    except(TypeError):

                        continue

                print(R+splitter+W)

                outWorkbook.close()

                userfdec = input(
                    C+'RD3 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()

            case 5:

                outWorkbook = xlsxwriter.Workbook(f"RD4_Daily_BL.xlsx")

                outSheet = outWorkbook.add_worksheet(name='Median_bl')

                # ==================== data
                df_RD4 = pd.read_excel('RD4_Daily_data.xlsx', sheet_name='Sheet0')
                astro_RD4 = df_RD4.to_dict('records')
                # ====================

                main_cell_source_index_4 = df_RD4[['cell_ref']].dropna()
                main_cell_source_index_4 = np.asanyarray(
                    main_cell_source_index_4).flatten()
                main_cell_source_index_4 = list(
                    np.nan_to_num(main_cell_source_index_4))

                kpi_list = [

                    
                    'total_traffic_volume(gb)',
                    'Level',
                    'e-rab_setup_success_rate(hu_cell)',
                    'e-rab_setup_success_rate',
                    'interf_hoout_sr',
                    'intraf_hoout_sr',
                    'average_ul_packet_loss_%(huawei_lte_ucell)',
                    'call_drop_rate',
                    'average_downlink_user_throughput(mbit/s)',
                    'average_uplink_user_throughput(mbit/s)',
                    'csfb_rate',
                    'cssr(all)',
                    'downlink_cell_throghput(kbit/s)',
                    'uplink_cell_throghput(kbit/s)',
                    'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell',
                    'rrc_connection_setup_success_rate_service',
                    's1signal_e-rab_setup_sr(hu_cell)'
      

                ]

                for z in range(len(main_cell_source_index_4)):

                    # ---------------------------------- 2G Data KPIs
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
              

                    for i in range(len(astro_RD4)):

                        if astro_RD4[i]['cell'] == main_cell_source_index_4[z]:

                            kpi_1.append(
                                astro_RD4[i]['total_traffic_volume(gb)'])
                            kpi_2.append(
                                astro_RD4[i]['e-rab_setup_success_rate(hu_cell)'])
                            kpi_3.append(
                                astro_RD4[i]['e-rab_setup_success_rate'])
                            kpi_4.append(astro_RD4[i]['interf_hoout_sr'])
                            kpi_5.append(astro_RD4[i]['intraf_hoout_sr'])
                            kpi_6.append(astro_RD4[i]['average_ul_packet_loss_%(huawei_lte_ucell)'])
                            kpi_7.append(
                                astro_RD4[i]['call_drop_rate'])
                            kpi_8.append(
                                astro_RD4[i]['average_downlink_user_throughput(mbit/s)'])
                            kpi_9.append(
                                astro_RD4[i]['average_uplink_user_throughput(mbit/s)'])
                            kpi_10.append(
                                astro_RD4[i]['csfb_rate'])
                            kpi_11.append(astro_RD4[i]['cssr(all)'])
                            kpi_12.append(
                                astro_RD4[i]['downlink_cell_throghput(kbit/s)'])
                            kpi_13.append(
                                astro_RD4[i]['uplink_cell_throghput(kbit/s)'])
                            kpi_14.append(astro_RD4[i]['intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell'])
                            kpi_15.append(astro_RD4[i]['rrc_connection_setup_success_rate_service'])
                            kpi_16.append(astro_RD4[i]['s1signal_e-rab_setup_sr(hu_cell)'])
                         

                        else:

                            continue
                    
                    row_1 = 0

                    row_2 = 0
                    column_2 = 1

                    for kpi in kpi_list:

                        outSheet.write(row_2, column_2, kpi)

                        column_2 += 1

                    outSheet.write(row_2, row_1, 'cell')

                    try:

                        print(f'{z} cell : {C+main_cell_source_index_4[z]+W}')
                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_4[z])

                        print(Y+'total_traffic_volume(gb)'+W, f'= {kpi_1}'+G,
                            f'Average : {float(np.nanmean(kpi_1))}'+W)    # edit from here 

                        outSheet.write(z + 1, 1 , float(np.nanmean(kpi_1)))


                        if float(np.nanmean(kpi_1)) > 160:

                            outSheet.write(                     
                            z + 1, 2, 'L1')   

                        elif 120 < float(np.nanmean(kpi_1)) <= 160:

                            outSheet.write(                     
                            z + 1, 2, 'L2')   
                        
                        elif 80 < float(np.nanmean(kpi_1)) <= 120:

                            outSheet.write(                     
                            z + 1, 2, 'L3')   

                        elif 40 < float(np.nanmean(kpi_1)) <= 80:         

                            outSheet.write(                     
                            z + 1, 2, 'L4')   

                        elif float(np.nanmean(kpi_1)) <= 40:

                            outSheet.write(                     
                            z + 1, 2, 'L5')


                        print(Y+'e-rab_setup_success_rate(hu_cell)'+W, f'= {kpi_2}'+G,
                            f'Median : {float(np.nanmedian(kpi_2))}'+W)

                        outSheet.write(z + 1, 3 , float(np.nanmedian(kpi_2)))

                        print(Y+'e-rab_setup_success_rate'+W, f'= {kpi_3}'+G,
                            f'Median : {float(np.nanmedian(kpi_3))}'+W)

                        outSheet.write(z + 1, 4 , float(np.nanmedian(kpi_3)))

                        print(Y+'interf_hoout_sr'+W, f'= {kpi_4}'+G,
                            f'Median : {float(np.nanmedian(kpi_4))}'+W)

                        outSheet.write(z + 1, 5 , float(np.nanmedian(kpi_4)))

                        print(Y+'intraf_hoout_sr'+W, f'= {kpi_5}'+G,
                            f'Median : {float(np.nanmedian(kpi_5))}'+W)

                        outSheet.write(z + 1, 6 , float(np.nanmedian(kpi_5)))

                        print(Y+'average_ul_packet_loss_%(huawei_lte_ucell)'+W, f'= {kpi_6}'+G,
                            f'Median : {float(np.nanmedian(kpi_6))}'+W)

                        outSheet.write(z + 1, 7 , float(np.nanmedian(kpi_6)))

                        print(Y+'call_drop_rate'+W, f'= {kpi_7}'+G,
                            f'Median : {float(np.nanmedian(kpi_7))}'+W)

                        outSheet.write(z + 1, 8 , float(np.nanmedian(kpi_7)))

                        print(Y+'average_downlink_user_throughput(mbit/s)'+W, f'= {kpi_8}'+G,
                            f'Median : {float(np.nanmedian(kpi_8))}'+W)

                        outSheet.write(z + 1, 9 , float(np.nanmedian(kpi_8)))

                        print(Y+'average_uplink_user_throughput(mbit/s)'+W, f'= {kpi_9}'+G,
                            f'Median : {float(np.nanmedian(kpi_9))}'+W)

                        outSheet.write(z + 1, 10 , float(np.nanmedian(kpi_9)))

                        print(Y+'csfb_rate'+W, f'= {kpi_10}'+G,
                            f'Median : {float(np.nanmedian(kpi_10))}'+W)

                        outSheet.write(z + 1, 11 , float(np.nanmedian(kpi_10)))

                        print(Y+'cssr(all)'+W, f'= {kpi_11}'+G,
                            f'Median : {float(np.nanmedian(kpi_11))}'+W)

                        outSheet.write(z + 1, 12 , float(np.nanmedian(kpi_11)))

                        print(Y+'downlink_cell_throghput(kbit/s)'+W, f'= {kpi_12}'+G,
                            f'Median : {float(np.nanmedian(kpi_12))}'+W)

                        outSheet.write(z + 1, 13 , float(np.nanmedian(kpi_12)))

                        print(Y+'uplink_cell_throghput(kbit/s)'+W, f'= {kpi_13}'+G,
                            f'Median : {float(np.nanmedian(kpi_13))}'+W)

                        outSheet.write(z + 1, 14 , float(np.nanmedian(kpi_13)))

                        print(Y+'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell'+W, f'= {kpi_14}'+G,
                            f'Median : {float(np.nanmedian(kpi_14))}'+W)

                        outSheet.write(z + 1, 15 , float(np.nanmedian(kpi_14)))

                        print(Y+'rrc_connection_setup_success_rate_service'+W, f'= {kpi_15}'+G,
                            f'Median : {float(np.nanmedian(kpi_15))}'+W)

                        outSheet.write(z + 1, 16 , float(np.nanmedian(kpi_15)))

                        print(Y+'s1signal_e-rab_setup_sr(hu_cell)'+W, f'= {kpi_16}'+G,
                            f'Median : {float(np.nanmedian(kpi_16))}'+W)

                        outSheet.write(z + 1, 17 , float(np.nanmedian(kpi_16)))
                        
                        
                        
                    except(TypeError):

                        continue


                print(R+splitter+W)

                outWorkbook.close()

                userfdec = input(
                    C+'RD4 calculation Done! continue? [y/n] : '+W)

                userfdec = userfdec.lower()

                match userfdec:

                    case 'y':

                        continue

                    case 'n':

                        sys.exit()

            case 6:

                os.chdir(gitdir)

                os.system(fr'python newcell.py')
