import pyfiglet
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

datadir = fr'{gitdir}\Data'

basedir = fr'{gitdir}\Daily_baselines'

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


        datamode = input('Select Daily or BH [d/b] : ')

        datamode = datamode.lower()

        match datamode:

            case 'd':

                beg_calc = input('Enter Calculation start date : ')
                end_calc = input('Enter Calculation end date : ')

                userCh = int(input('Select tech : '))

                        
                match userCh:

                    case 1:


                        outWorkbook = xlsxwriter.Workbook(f"CC2_Daily_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='2G_VOICE_DLY')

                        # ================= data
                        df_CC2 = pd.read_excel('CC2_Daily_data.xlsx', sheet_name='Sheet1')
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



                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_2[z][0:2]))



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
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

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
                            C+'CC2-Daily calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 2:

                        # os.makedirs(fr'{gitdir}/RD2_BL')
                        outWorkbook = xlsxwriter.Workbook(f"RD2_Daily_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='2G_DATA_DLY')

                        # ================= data
                        df_RD2 = pd.read_excel('RD2_Daily_data.xlsx', sheet_name='Sheet1')
                        astro_RD2 = df_RD2.to_dict('records')
                        # ========================

                        main_cell_source_index_2 = df_RD2[['cell_ref']].dropna()
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
                            'Payload_Total',
                            'TBF_Establishment_Success_Rate(UL+DL)(%)',
                            'TBF_Drop(UL+DL)',
                            'Average_Throughput_of_Downlink_GPRS_LLC_per_User(kbps)',
                            'Average_Throughput_of_Downlink_EGPRS_LLC_per_User(kbps)',
                            'THR_DL_GPRS_PER_TS',
                            'THR_DL_EGPRS_PER_TS',
                            'Edge_share_Payload',
                            'Level'

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
                            column_2 = 0

                            for kpi in kpi_list:

                                outSheet.write(row_2, column_2, kpi)

                                column_2 += 1

                        

                            try:

                                

                                print(f'{z} cell : {C+main_cell_source_index_2[z]+W}')
                                outSheet.write(
                                    z + 1, 4, main_cell_source_index_2[z])


                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_2[z][0:2]))
                                
                                

                                print(Y+'payload_total(cell_hu)'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)

                                outSheet.write(
                                    z + 1, 5, float(np.nanmean(kpi_1)))

                                if float(np.nanmean(kpi_1)) > 1:

                                    outSheet.write(                     
                                    z + 1, 13, 'L1')   

                                elif 0.7 < float(np.nanmean(kpi_1)) <= 1:

                                    outSheet.write(                     
                                    z + 1, 13, 'L2')   
                                
                                elif 0.4 < float(np.nanmean(kpi_1)) <= 0.7:

                                    outSheet.write(                     
                                    z + 1, 13, 'L3')   

                                elif 0.2 < float(np.nanmean(kpi_1)) <= 0.4:         

                                    outSheet.write(                     
                                    z + 1, 13, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 0.2:

                                    outSheet.write(                     
                                    z + 1, 13, 'L5')

                                print(Y+'tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_2)))




                                print(Y+'tbf_drop(ul+dl)(hu_cell)'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_3)))





                                print(Y+'average_throughput_of_downlink_gprs_llc_per_user(kbps)'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_4)))




                                print(Y+'average_throughput_of_downlink_egprs_llc_per_user(kbps)'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_5)))






                                print(Y+'thr_dl_gprs_per_ts(cell_hu)'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_6)))






                                print(Y+'thr_dl_egprs_per_ts(cell_hu)'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_7)))



                                

                                print(Y+'edge_share_payload(cell_hu)'+W, f'= {kpi_8}'+G,
                                    f'Median : {float(np.nanmedian(kpi_8))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_8))) == True:
                                    outSheet.write(
                                    z + 1, 12, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 12, float(np.nanmedian(kpi_8)))


                

                            except(TypeError):

                                continue

                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'RD2-Daily calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 3:

                        outWorkbook = xlsxwriter.Workbook(f"CC3_Daily_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='3G_VOICE_DLY')

                        # ==================== data
                        df_CC3 = pd.read_excel('CC3_Daily_data.xlsx', sheet_name='Sheet1')
                        astro_CC3 = df_CC3.to_dict('records')
                        # ====================

                        main_cell_source_index_3 = df_CC3[['cell_ref']].dropna()
                        main_cell_source_index_3 = np.asanyarray(
                            main_cell_source_index_3).flatten()
                        main_cell_source_index_3 = list(
                            np.nan_to_num(main_cell_source_index_3))

                        kpi_list = [


                            'Calculation Period',
                            'Region',
                            'Province',
                            'RNC',
                            'Cell',
                            'cs_erlang',
                            'CS_RAB_Setup_Success_Ratio',
                            'CS_IRAT_HO_SR',
                            'InterFrequency_Hardhandover_success_Ratio_CSservice',
                            'AMR_Call_Drop_Ratio',
                            'Softer_Handover_Success_Ratio',
                            'CS_RRC_Connection_Establishment_SR',
                            'Level',

                            

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
                                        astro_CC3[i]['cs_irat_ho_sr'])
                                    kpi_4.append(
                                        astro_CC3[i]['interfrequency_hardhandover_success_ratio_csservice'])
                                    kpi_5.append(
                                        astro_CC3[i]['amr_call_drop_ratio_new(hu_cell)'])
                                    kpi_6.append(
                                        astro_CC3[i]['softer_handover_success_ratio(hu_cell)'])
                                    kpi_7.append(
                                        astro_CC3[i]['cs_rrc_connection_establishment_sr'])
                                    

                                else:

                                    continue

                            row_1 = 0

                            row_2 = 0
                            column_2 = 0

                            for kpi in kpi_list:

                                outSheet.write(row_2, column_2, kpi)

                                column_2 += 1

                    

                            try:

                                print(f'{z} cell : {C+main_cell_source_index_3[z]+W}')
                                outSheet.write(
                                    z + 1, 4, main_cell_source_index_3[z])



                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_3[z][0:2]))



                                print(Y+'cs_erlang'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)

                                outSheet.write(
                                    z + 1, 5, float(np.nanmean(kpi_1)))

                                if float(np.nanmean(kpi_1)) > 150:

                                    outSheet.write(                     
                                    z + 1, 12, 'L1')   

                                elif 120 < float(np.nanmean(kpi_1)) <= 150:

                                    outSheet.write(                     
                                    z + 1, 12, 'L2')   
                                
                                elif 90 < float(np.nanmean(kpi_1)) <= 120:

                                    outSheet.write(                     
                                    z + 1, 12, 'L3')   

                                elif 50 < float(np.nanmean(kpi_1)) <= 90:         

                                    outSheet.write(                     
                                    z + 1, 12, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 50:

                                    outSheet.write(                     
                                    z + 1, 12, 'L5')

                                print(Y+'cs_rab_setup_success_ratio'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_2)))




                                print(Y+'cs_irat_ho_sr'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_3)))




                                print(Y+'interfrequency_hardhandover_success_ratio_csservice'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_4)))




                                print(Y+'amr_call_drop_ratio_new(hu_cell)'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_5)))





                                print(Y+'softer_handover_success_ratio(hu_cell)'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_6)))




                                print(Y+'cs_rrc_connection_establishment_sr'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_7)))
                                



                            except(TypeError):

                                continue

                            

                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'CC3-Daily calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 4:

                        outWorkbook = xlsxwriter.Workbook(f"RD3_Daily_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='3G_DATA_DLY')

                        # ==================== data
                        df_RD3 = pd.read_excel('RD3_Daily_data.xlsx', sheet_name='Sheet1')
                        astro_RD3 = df_RD3.to_dict('records')
                        # ====================

                        main_cell_source_index_3 = df_RD3[['cell_ref']].dropna()
                        main_cell_source_index_3 = np.asanyarray(
                            main_cell_source_index_3).flatten()
                        main_cell_source_index_3 = list(
                            np.nan_to_num(main_cell_source_index_3))

                        kpi_list = [
                            
                            'Calculation Period',
                            'Region',
                            'Province',
                            'RNC',
                            'Cell',
                            'Payload',
                            'AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)',
                            'CS+PS_RAB_Setup_Success_Ratio',
                            'HSDPA_Soft_HandOver_Success_Ratio',
                            'HS_share_PAYLOAD_%',
                            'HSDPA_cdr(%)',
                            'HSDPA_SCHEDULING_Cell_throughput',
                            'PS_RAB_Setup_Success_Ratio',
                            'PS_RRC_Connection_success_Rate_repeatless',
                            'PS_Call_Drop_Ratio',
                            'PS_CSSR',
                            'hsupa_uplink_throughput_in_V16',
                            'hsdpa_rab_setup_success_ratio(hu_cell)',
                            'hsupa_rab_setup_success_ratio(hu_cell)',
                            'hsupa_cdr(%)_(hu_cell)_new',
                            'Level'



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
                                    kpi_2.append(astro_RD3[i]['average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'])
                                    kpi_3.append(astro_RD3[i]['cs+ps_rab_setup_success_ratio'])
                                    kpi_4.append(
                                        astro_RD3[i]['hsdpa_soft_handover_success_ratio'])
                                    kpi_5.append(
                                        astro_RD3[i]['hs_share_payload_%'])
                                    kpi_6.append(
                                        astro_RD3[i]['hsdpa_cdr(%)_(hu_cell)_new'])
                                    kpi_7.append(
                                        astro_RD3[i]['hsdpa_scheduling_cell_throughput(cell_huawei)'])
                                    kpi_8.append(astro_RD3[i]['ps_rab_setup_success_ratio'])
                                    kpi_9.append(
                                        astro_RD3[i]['ps_rrc_connection_success_rate_repeatless(hu_cell)'])
                                    kpi_10.append(
                                        astro_RD3[i]['ps_call_drop_ratio'])
                                    kpi_11.append(
                                        astro_RD3[i]['ps_cssr'])
                                    kpi_12.append(
                                        astro_RD3[i]['hsupa_uplink_throughput_in_v16(cell_hu)'])
                                    kpi_13.append(
                                        astro_RD3[i]['hsdpa_rab_setup_success_ratio(hu_cell)'])
                                    kpi_14.append(
                                        astro_RD3[i]['hsupa_rab_setup_success_ratio(hu_cell)'])
                                    kpi_15.append(
                                        astro_RD3[i]['hsupa_cdr(%)_(hu_cell)_new'])
                                

                                else:

                                    continue

                            row_1 = 0

                            row_2 = 0
                            column_2 = 0

                            for kpi in kpi_list:

                                outSheet.write(row_2, column_2, kpi)

                                column_2 += 1



                            try:

                                print(f'{z} cell : {C+main_cell_source_index_3[z]+W}')
                                outSheet.write(
                                    z + 1, 4, main_cell_source_index_3[z])

                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_3[z][0:2]))


                                print(Y+'payload'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)

                                outSheet.write(
                                    z + 1, 5, float(np.nanmean(kpi_1)))

                                
                                if float(np.nanmean(kpi_1)) > 28:

                                    outSheet.write(                     
                                    z + 1, 20, 'L1')   

                                elif 23 < float(np.nanmean(kpi_1)) <= 28:

                                    outSheet.write(                     
                                    z + 1, 20, 'L2')   
                                
                                elif 18 < float(np.nanmean(kpi_1)) <= 23:

                                    outSheet.write(                     
                                    z + 1, 20, 'L3')   

                                elif 13 < float(np.nanmean(kpi_1)) <= 18:         

                                    outSheet.write(                     
                                    z + 1, 20, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 13:

                                    outSheet.write(                     
                                    z + 1, 20, 'L5')

                                print(Y+'average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_2)))



            

                                print(Y+'cs+ps_rab_setup_success_ratio'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_3)))





                                print(Y+'hsdpa_soft_handover_success_ratio'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_4)))




                                print(Y+'hs_share_payload_%'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_5)))




                                print(Y+'hsdpa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_6)))




                                print(Y+'hsdpa_scheduling_cell_throughput(cell_huawei)'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_7)))




                                print(Y+'ps_rab_setup_success_ratio'+W, f'= {kpi_8}'+G,
                                    f'Median : {float(np.nanmedian(kpi_8))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_8))) == True:
                                    outSheet.write(
                                    z + 1, 12, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 12, float(np.nanmedian(kpi_8)))





                                print(Y+'ps_rrc_connection_success_rate_repeatless(hu_cell)'+W, f'= {kpi_9}'+G,
                                    f'Median : {float(np.nanmedian(kpi_9))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_9))) == True:
                                    outSheet.write(
                                    z + 1, 13, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 13, float(np.nanmedian(kpi_9)))




                                print(Y+'ps_call_drop_ratio'+W, f'= {kpi_10}'+G,
                                    f'Median : {float(np.nanmedian(kpi_10))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_10))) == True:
                                    outSheet.write(
                                    z + 1, 14, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 14, float(np.nanmedian(kpi_10)))




                                print(Y+'ps_cssr'+W, f'= {kpi_11}'+G,
                                    f'Median : {float(np.nanmedian(kpi_11))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_11))) == True:
                                    outSheet.write(
                                    z + 1, 15, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 15, float(np.nanmedian(kpi_11)))




                                print(Y+'hsupa_uplink_throughput_in_v16(cell_hu)'+W, f'= {kpi_12}'+G,
                                    f'Median : {float(np.nanmedian(kpi_12))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_12))) == True:
                                    outSheet.write(
                                    z + 1, 16, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 16, float(np.nanmedian(kpi_12)))




                                print(Y+'hsdpa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_13}'+G,
                                    f'Median : {float(np.nanmedian(kpi_13))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_13))) == True:
                                    outSheet.write(
                                    z + 1, 17, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 17, float(np.nanmedian(kpi_13)))





                                print(Y+'hsupa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_14}'+G,
                                    f'Median : {float(np.nanmedian(kpi_14))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_14))) == True:
                                    outSheet.write(
                                    z + 1, 18, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 18, float(np.nanmedian(kpi_14)))





                                print(Y+'hsupa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_15}'+G,
                                    f'Median : {float(np.nanmedian(kpi_15))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_15))) == True:
                                    outSheet.write(
                                    z + 1, 19, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 19, float(np.nanmedian(kpi_15)))

                            


                            except(TypeError):

                                continue

                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'RD3-Daily calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 5:

                        outWorkbook = xlsxwriter.Workbook(f"RD4_Daily_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='4G_DATA_DLY')

                        # ==================== data
                        df_RD4 = pd.read_excel('RD4_Daily_data.xlsx', sheet_name='Sheet1')
                        astro_RD4 = df_RD4.to_dict('records')
                        # ====================

                        main_cell_source_index_4 = df_RD4[['cell_ref']].dropna()
                        main_cell_source_index_4 = np.asanyarray(
                            main_cell_source_index_4).flatten()
                        main_cell_source_index_4 = list(
                            np.nan_to_num(main_cell_source_index_4))

                        kpi_list = [

                            
                            'Calculation Period',
                            'Region',
                            'Province',
                            'Cell',
                            'total_traffic_volume(gb)',
                            'Average_Downlink_User_Throughput(Mbit/s)',
                            'Call_Drop_Rate',
                            'Average_UPlink_User_Throughput(Mbit/s)',
                            'RRC_Connection_Setup_Success_Rate_service',
                            'E-RAB_Setup_Success_Rate',
                            'E-RAB_Setup_Success_Rate(Hu_Cell)',
                            'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell)',
                            'CSFB_Rate',
                            'S1Signal_E-RAB_Setup_SR(Hu_Cell)',
                            'InterF_HOOut_SR',
                            'IntraF_HOOut_SR',
                            'downlink_cell_throghput(kbit/s)',
                            'uplink_cell_throghput(kbit/s)',
                            'Average_UL_Packet_Loss_%(Huawei_LTE_UCell)',
                            'CSSR(ALL)',
                            'Level'

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
                                        astro_RD4[i]['average_downlink_user_throughput(mbit/s)'])
                                    kpi_3.append(
                                        astro_RD4[i]['call_drop_rate'])
                                    kpi_4.append(astro_RD4[i]['average_uplink_user_throughput(mbit/s)'])
                                    kpi_5.append(astro_RD4[i]['rrc_connection_setup_success_rate_service'])
                                    kpi_6.append(astro_RD4[i]['e-rab_setup_success_rate'])
                                    kpi_7.append(
                                        astro_RD4[i]['e-rab_setup_success_rate(hu_cell)'])
                                    kpi_8.append(
                                        astro_RD4[i]['intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell)'])
                                    kpi_9.append(
                                        astro_RD4[i]['csfb_rate'])
                                    kpi_10.append(
                                        astro_RD4[i]['s1signal_e-rab_setup_sr(hu_cell)'])
                                    kpi_11.append(astro_RD4[i]['interf_hoout_sr'])
                                    kpi_12.append(
                                        astro_RD4[i]['intraf_hoout_sr'])
                                    kpi_13.append(
                                        astro_RD4[i]['downlink_cell_throghput(kbit/s)'])
                                    kpi_14.append(astro_RD4[i]['uplink_cell_throghput(kbit/s)'])
                                    kpi_15.append(astro_RD4[i]['average_ul_packet_loss_%(huawei_lte_ucell)'])
                                    kpi_16.append(astro_RD4[i]['cssr(all)'])
                                

                                else:

                                    continue
                            
                            row_1 = 0

                            row_2 = 0
                            column_2 = 0

                            for kpi in kpi_list:

                                outSheet.write(row_2, column_2, kpi)

                                column_2 += 1



                            try:

                                print(f'{z} cell : {C+main_cell_source_index_4[z]+W}')
                                outSheet.write(
                                    z + 1, 3, main_cell_source_index_4[z])


                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_4[z][0:2]))

                                    

                                print(Y+'total_traffic_volume(gb)'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)    # edit from here 

                                outSheet.write(z + 1, 4 , float(np.nanmean(kpi_1)))


                                if float(np.nanmean(kpi_1)) > 160:

                                    outSheet.write(                     
                                    z + 1, 20, 'L1')   

                                elif 120 < float(np.nanmean(kpi_1)) <= 160:

                                    outSheet.write(                     
                                    z + 1, 20, 'L2')   
                                
                                elif 80 < float(np.nanmean(kpi_1)) <= 120:

                                    outSheet.write(                     
                                    z + 1, 20, 'L3')   

                                elif 40 < float(np.nanmean(kpi_1)) <= 80:         

                                    outSheet.write(                     
                                    z + 1, 20, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 40:

                                    outSheet.write(                     
                                    z + 1, 20, 'L5')


                                print(Y+'average_downlink_user_throughput(mbit/s)'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 5, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 5, float(np.nanmedian(kpi_2)))


                                

                                print(Y+'Call_Drop_Rate'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_3)))





                                print(Y+'Average_UPlink_User_Throughput(Mbit/s)'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_4)))





                                print(Y+'RRC_Connection_Setup_Success_Rate_service'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_5)))




                                print(Y+'E-RAB_Setup_Success_Rate'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_6)))




                                print(Y+'E-RAB_Setup_Success_Rate(Hu_Cell)'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_7)))




                                print(Y+'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell'+W, f'= {kpi_8}'+G,
                                    f'Median : {float(np.nanmedian(kpi_8))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_8))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_8)))



                                print(Y+'CSFB_Rate'+W, f'= {kpi_9}'+G,
                                    f'Median : {float(np.nanmedian(kpi_9))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_9))) == True:
                                    outSheet.write(
                                    z + 1, 12, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 12, float(np.nanmedian(kpi_9)))




                                print(Y+'S1Signal_E-RAB_Setup_SR(Hu_Cell)'+W, f'= {kpi_10}'+G,
                                    f'Median : {float(np.nanmedian(kpi_10))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_10))) == True:
                                    outSheet.write(
                                    z + 1, 13, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 13, float(np.nanmedian(kpi_10)))



                                print(Y+'InterF_HOOut_SR'+W, f'= {kpi_11}'+G,
                                    f'Median : {float(np.nanmedian(kpi_11))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_11))) == True:
                                    outSheet.write(
                                    z + 1, 14, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 14, float(np.nanmedian(kpi_11)))




                                print(Y+'IntraF_HOOut_SR'+W, f'= {kpi_12}'+G,
                                    f'Median : {float(np.nanmedian(kpi_12))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_12))) == True:
                                    outSheet.write(
                                    z + 1, 15, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 15, float(np.nanmedian(kpi_12)))




                                print(Y+'downlink_cell_throghput(kbit/s)'+W, f'= {kpi_13}'+G,
                                    f'Median : {float(np.nanmedian(kpi_13))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_13))) == True:
                                    outSheet.write(
                                    z + 1, 16, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 16, float(np.nanmedian(kpi_13)))





                                print(Y+'uplink_cell_throghput(kbit/s)'+W, f'= {kpi_14}'+G,
                                    f'Median : {float(np.nanmedian(kpi_14))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_14))) == True:
                                    outSheet.write(
                                    z + 1, 17, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 17, float(np.nanmedian(kpi_14)))




                                print(Y+'Average_UL_Packet_Loss_%(Huawei_LTE_UCell)'+W, f'= {kpi_15}'+G,
                                    f'Median : {float(np.nanmedian(kpi_15))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_15))) == True:
                                    outSheet.write(
                                    z + 1, 18, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 18, float(np.nanmedian(kpi_15)))



                                print(Y+'CSSR(ALL)'+W, f'= {kpi_16}'+G,
                                    f'Median : {float(np.nanmedian(kpi_16))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_16))) == True:
                                    outSheet.write(
                                    z + 1, 19, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 19, float(np.nanmedian(kpi_16)))
                                
                                
                                
                            except(TypeError):

                                continue


                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'RD4-Daily calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 6:

                        os.chdir(gitdir)

                        os.system(fr'python newcell.py')


            case 'b':

                beg_calc = input('Enter Calculation start date : ')
                end_calc = input('Enter Calculation end date : ')

                userCh = int(input('Select tech : '))

                match userCh:

                    case 1:

                        

                        outWorkbook = xlsxwriter.Workbook(f"CC2_BH_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='2G_VOICE_BH')

                        # ================= data
                        df_CC2 = pd.read_excel('CC2_BH_data.xlsx', sheet_name='Sheet1')
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



                                if astro_CC2[i]['CELL'] == main_cell_source_index_2[z]:

                                    kpi_1.append(astro_CC2[i]['TCH_Traffic_BH'])
                                    kpi_2.append(astro_CC2[i]['CSSR3'])
                                    kpi_3.append(astro_CC2[i]['AMRHR_USAGE'])
                                    kpi_4.append(astro_CC2[i]['SDCCH_Congestion_Rate'])
                                    kpi_5.append(astro_CC2[i]['SDCCH_Drop_Rate'])
                                    kpi_6.append(astro_CC2[i]['TCH_Assignment_FR'])
                                    kpi_7.append(astro_CC2[i]['IHSR2'])
                                    kpi_8.append(astro_CC2[i]['OHSR2'])
                                    kpi_9.append(astro_CC2[i]['SDCCH_Access_Success_Rate2'])
                                    kpi_10.append(astro_CC2[i]['CDR3'])
                                    kpi_11.append(astro_CC2[i]['RX_QUALITTY_DL_NEW'])
                                    kpi_12.append(astro_CC2[i]['RX_QUALITTY_UL_NEW'])
                        

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

                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_2[z][0:2]))



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
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

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
                            C+'CC2-BH calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 2:

                        
                        # os.makedirs(fr'{gitdir}/RD2_BL')
                        outWorkbook = xlsxwriter.Workbook(f"RD2_BH_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='2G_DATA_BH')

                        # ================= data
                        df_RD2 = pd.read_excel('RD2_BH_data.xlsx', sheet_name='Sheet1')
                        astro_RD2 = df_RD2.to_dict('records')
                        # ========================

                        main_cell_source_index_2 = df_RD2[['cell_ref']].dropna()
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
                            'Payload_Total',
                            'TBF_Establishment_Success_Rate(UL+DL)(%)',
                            'TBF_Drop(UL+DL)',
                            'Average_Throughput_of_Downlink_GPRS_LLC_per_User(kbps)',
                            'Average_Throughput_of_Downlink_EGPRS_LLC_per_User(kbps)',
                            'THR_DL_GPRS_PER_TS',
                            'THR_DL_EGPRS_PER_TS',
                            'Edge_share_Payload',
                            'Level'

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

                                if astro_RD2[i]['CELL'] == main_cell_source_index_2[z]:

                                    kpi_1.append(
                                        astro_RD2[i]['Payload_Total(CELL_HU)'])
                                    kpi_2.append(
                                        astro_RD2[i]['TBF_Establishment_Success_Rate(UL+DL)(%)(HU_Cell)'])
                                    kpi_3.append(
                                        astro_RD2[i]['TBF_Drop(UL+DL)(HU_Cell)'])
                                    kpi_4.append(
                                        astro_RD2[i]['Average_Throughput_of_Downlink_GPRS_LLC_per_User(kbps)'])
                                    kpi_5.append(
                                        astro_RD2[i]['Average_Throughput_of_Downlink_EGPRS_LLC_per_User(kbps)'])
                                    kpi_6.append(
                                        astro_RD2[i]['THR_DL_GPRS_PER_TS(CELL_HU)'])
                                    kpi_7.append(
                                        astro_RD2[i]['THR_DL_EGPRS_PER_TS(CELL_HU)'])
                                    kpi_8.append(
                                        astro_RD2[i]['Edge_share_Payload(CELL_HU)'])
                                

                                else:

                                    continue

                            # ================================== excel writing main
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


                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_2[z][0:2]))

                                
                                

                                print(Y+'payload_total(cell_hu)'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)

                                outSheet.write(
                                    z + 1, 5, float(np.nanmean(kpi_1)))

                                if float(np.nanmean(kpi_1)) > 1:

                                    outSheet.write(                     
                                    z + 1, 13, 'L1')   

                                elif 0.7 < float(np.nanmean(kpi_1)) <= 1:

                                    outSheet.write(                     
                                    z + 1, 13, 'L2')   
                                
                                elif 0.4 < float(np.nanmean(kpi_1)) <= 0.7:

                                    outSheet.write(                     
                                    z + 1, 13, 'L3')   

                                elif 0.2 < float(np.nanmean(kpi_1)) <= 0.4:         

                                    outSheet.write(                     
                                    z + 1, 13, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 0.2:

                                    outSheet.write(                     
                                    z + 1, 13, 'L5')

                                print(Y+'tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_2)))




                                print(Y+'tbf_drop(ul+dl)(hu_cell)'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_3)))





                                print(Y+'average_throughput_of_downlink_gprs_llc_per_user(kbps)'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_4)))




                                print(Y+'average_throughput_of_downlink_egprs_llc_per_user(kbps)'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_5)))






                                print(Y+'thr_dl_gprs_per_ts(cell_hu)'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_6)))






                                print(Y+'thr_dl_egprs_per_ts(cell_hu)'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_7)))



                                

                                print(Y+'edge_share_payload(cell_hu)'+W, f'= {kpi_8}'+G,
                                    f'Median : {float(np.nanmedian(kpi_8))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_8))) == True:
                                    outSheet.write(
                                    z + 1, 12, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 12, float(np.nanmedian(kpi_8)))


                

                            except(TypeError):

                                continue

                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'RD2-BH calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 3:

                        outWorkbook = xlsxwriter.Workbook(f"CC3_BH_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='3G_VOICE_BH')

                        # ==================== data
                        df_CC3 = pd.read_excel('CC3_BH_data.xlsx', sheet_name='Sheet1')
                        astro_CC3 = df_CC3.to_dict('records')
                        # ====================

                        main_cell_source_index_3 = df_CC3[['cell_ref']].dropna()
                        main_cell_source_index_3 = np.asanyarray(
                            main_cell_source_index_3).flatten()
                        main_cell_source_index_3 = list(
                            np.nan_to_num(main_cell_source_index_3))

                        kpi_list = [


                            'Calculation Period',
                            'Region',
                            'Province',
                            'RNC',
                            'Cell',
                            'cs_erlang',
                            'CS_RAB_Setup_Success_Ratio',
                            'CS_IRAT_HO_SR',
                            'InterFrequency_Hardhandover_success_Ratio_CSservice',
                            'AMR_Call_Drop_Ratio',
                            'Softer_Handover_Success_Ratio',
                            'CS_RRC_Connection_Establishment_SR',
                            'Level',

                            

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

                                if astro_CC3[i]['CELL'] == main_cell_source_index_3[z]:

                                    kpi_1.append(astro_CC3[i]['CS_TrafficBH'])
                                    kpi_2.append(
                                        astro_CC3[i]['CS_RAB_Setup_Success_Ratio'])
                                    kpi_3.append(
                                        astro_CC3[i]['CS_IRAT_HO_SR'])
                                    kpi_4.append(
                                        astro_CC3[i]['InterFrequency_Hardhandover_success_Ratio_CSservice'])
                                    kpi_5.append(
                                        astro_CC3[i]['AMR_Call_Drop_Ratio_New(Hu_CELL)'])
                                    kpi_6.append(
                                        astro_CC3[i]['Softer_Handover_Success_Ratio(Hu_Cell)'])
                                    kpi_7.append(
                                        astro_CC3[i]['CS_RRC_Connection_Establishment_SR'])
                                    

                                else:

                                    continue

                            row_1 = 0

                            row_2 = 0
                            column_2 = 0

                            for kpi in kpi_list:

                                outSheet.write(row_2, column_2, kpi)

                                column_2 += 1

                    

                            try:

                                print(f'{z} cell : {C+main_cell_source_index_3[z]+W}')
                                outSheet.write(
                                    z + 1, 4, main_cell_source_index_3[z])

                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_3[z][0:2]))



                                print(Y+'cs_erlang'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)

                                outSheet.write(
                                    z + 1, 5, float(np.nanmean(kpi_1)))

                                if float(np.nanmean(kpi_1)) > 150:

                                    outSheet.write(                     
                                    z + 1, 12, 'L1')   

                                elif 120 < float(np.nanmean(kpi_1)) <= 150:

                                    outSheet.write(                     
                                    z + 1, 12, 'L2')   
                                
                                elif 90 < float(np.nanmean(kpi_1)) <= 120:

                                    outSheet.write(                     
                                    z + 1, 12, 'L3')   

                                elif 50 < float(np.nanmean(kpi_1)) <= 90:         

                                    outSheet.write(                     
                                    z + 1, 12, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 50:

                                    outSheet.write(                     
                                    z + 1, 12, 'L5')

                                print(Y+'cs_rab_setup_success_ratio'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_2)))




                                print(Y+'cs_irat_ho_sr'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_3)))




                                print(Y+'interfrequency_hardhandover_success_ratio_csservice'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_4)))




                                print(Y+'amr_call_drop_ratio_new(hu_cell)'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_5)))





                                print(Y+'softer_handover_success_ratio(hu_cell)'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_6)))




                                print(Y+'cs_rrc_connection_establishment_sr'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_7)))
                                



                            except(TypeError):

                                continue

                            

                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'CC3-BH calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 4:

                        outWorkbook = xlsxwriter.Workbook(f"RD3_BH_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='3G_DATA_BH')

                        # ==================== data
                        df_RD3 = pd.read_excel('RD3_BH_data.xlsx', sheet_name='Sheet1')
                        astro_RD3 = df_RD3.to_dict('records')
                        # ====================

                        main_cell_source_index_3 = df_RD3[['cell_ref']].dropna()
                        main_cell_source_index_3 = np.asanyarray(
                            main_cell_source_index_3).flatten()
                        main_cell_source_index_3 = list(
                            np.nan_to_num(main_cell_source_index_3))

                        kpi_list = [

                            
                            'Calculation Period',
                            'Region',
                            'Province',
                            'RNC',
                            'Cell',
                            'Payload',
                            'AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)',
                            'CS+PS_RAB_Setup_Success_Ratio',
                            'HSDPA_Soft_HandOver_Success_Ratio',
                            'HS_share_PAYLOAD_%',
                            'HSDPA_cdr(%)',
                            'HSDPA_SCHEDULING_Cell_throughput',
                            'PS_RAB_Setup_Success_Ratio',
                            'PS_RRC_Connection_success_Rate_repeatless',
                            'PS_Call_Drop_Ratio',
                            'PS_CSSR',
                            'hsupa_uplink_throughput_in_V16',
                            'hsdpa_rab_setup_success_ratio(hu_cell)',
                            'hsupa_rab_setup_success_ratio(hu_cell)',
                            'hsupa_cdr(%)_(hu_cell)_new',
                            'Level'


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

                                if astro_RD3[i]['CELL'] == main_cell_source_index_3[z]:

                                    kpi_1.append(astro_RD3[i]['Payload_Total_BH'])
                                    kpi_2.append(astro_RD3[i]['AVERAGE_HSDPA_USER_THROUGHPUT_DC+SC(Mbit/s)(CELL_HUAWEI)'])
                                    kpi_3.append(astro_RD3[i]['CS+PS_RAB_Setup_Success_Ratio'])
                                    kpi_4.append(
                                        astro_RD3[i]['HSDPA_Soft_HandOver_Success_Ratio'])
                                    kpi_5.append(
                                        astro_RD3[i]['HS_share_PAYLOAD_%'])
                                    kpi_6.append(
                                        astro_RD3[i]['HSDPA_cdr(%)_(Hu_Cell)_new'])
                                    kpi_7.append(
                                        astro_RD3[i]['HSDPA_SCHEDULING_Cell_throughput(CELL_HUAWEI)'])
                                    kpi_8.append(astro_RD3[i]['PS_RAB_Setup_Success_Ratio'])
                                    kpi_9.append(
                                        astro_RD3[i]['PS_RRC_Connection_success_Rate_repeatless(Hu_Cell)'])
                                    kpi_10.append(
                                        astro_RD3[i]['PS_Call_Drop_Ratio'])
                                    kpi_11.append(
                                        astro_RD3[i]['PS_CSSR'])
                                    kpi_12.append(
                                        astro_RD3[i]['hsupa_uplink_throughput_in_V16(CELL_Hu)'])
                                    kpi_13.append(
                                        astro_RD3[i]['HSDPA_RAB_Setup_Success_Ratio(Hu_Cell)'])
                                    kpi_14.append(
                                        astro_RD3[i]['HSUPA_RAB_Setup_Success_Ratio(Hu_Cell)'])
                                    kpi_15.append(
                                        astro_RD3[i]['HSUPA_CDR(%)_(Hu_Cell)_new'])
                                

                                else:

                                    continue

                            row_1 = 0

                            row_2 = 0
                            column_2 = 0

                            for kpi in kpi_list:

                                outSheet.write(row_2, column_2, kpi)

                                column_2 += 1



                            try:

                                print(f'{z} cell : {C+main_cell_source_index_3[z]+W}')
                                outSheet.write(
                                    z + 1, 4, main_cell_source_index_3[z])


                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_3[z][0:2]))



                                print(Y+'payload'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)

                                outSheet.write(
                                    z + 1, 5, float(np.nanmean(kpi_1)))

                                
                                if float(np.nanmean(kpi_1)) > 28:

                                    outSheet.write(                     
                                    z + 1, 20, 'L1')   

                                elif 23 < float(np.nanmean(kpi_1)) <= 28:

                                    outSheet.write(                     
                                    z + 1, 20, 'L2')   
                                
                                elif 18 < float(np.nanmean(kpi_1)) <= 23:

                                    outSheet.write(                     
                                    z + 1, 20, 'L3')   

                                elif 13 < float(np.nanmean(kpi_1)) <= 18:         

                                    outSheet.write(                     
                                    z + 1, 20, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 13:

                                    outSheet.write(                     
                                    z + 1, 20, 'L5')

                                print(Y+'average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_2)))



            

                                print(Y+'cs+ps_rab_setup_success_ratio'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_3)))





                                print(Y+'hsdpa_soft_handover_success_ratio'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_4)))




                                print(Y+'hs_share_payload_%'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_5)))




                                print(Y+'hsdpa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_6)))




                                print(Y+'hsdpa_scheduling_cell_throughput(cell_huawei)'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_7)))




                                print(Y+'ps_rab_setup_success_ratio'+W, f'= {kpi_8}'+G,
                                    f'Median : {float(np.nanmedian(kpi_8))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_8))) == True:
                                    outSheet.write(
                                    z + 1, 12, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 12, float(np.nanmedian(kpi_8)))





                                print(Y+'ps_rrc_connection_success_rate_repeatless(hu_cell)'+W, f'= {kpi_9}'+G,
                                    f'Median : {float(np.nanmedian(kpi_9))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_9))) == True:
                                    outSheet.write(
                                    z + 1, 13, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 13, float(np.nanmedian(kpi_9)))




                                print(Y+'ps_call_drop_ratio'+W, f'= {kpi_10}'+G,
                                    f'Median : {float(np.nanmedian(kpi_10))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_10))) == True:
                                    outSheet.write(
                                    z + 1, 14, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 14, float(np.nanmedian(kpi_10)))




                                print(Y+'ps_cssr'+W, f'= {kpi_11}'+G,
                                    f'Median : {float(np.nanmedian(kpi_11))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_11))) == True:
                                    outSheet.write(
                                    z + 1, 15, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 15, float(np.nanmedian(kpi_11)))




                                print(Y+'hsupa_uplink_throughput_in_v16(cell_hu)'+W, f'= {kpi_12}'+G,
                                    f'Median : {float(np.nanmedian(kpi_12))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_12))) == True:
                                    outSheet.write(
                                    z + 1, 16, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 16, float(np.nanmedian(kpi_12)))




                                print(Y+'hsdpa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_13}'+G,
                                    f'Median : {float(np.nanmedian(kpi_13))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_13))) == True:
                                    outSheet.write(
                                    z + 1, 17, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 17, float(np.nanmedian(kpi_13)))





                                print(Y+'hsupa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_14}'+G,
                                    f'Median : {float(np.nanmedian(kpi_14))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_14))) == True:
                                    outSheet.write(
                                    z + 1, 18, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 18, float(np.nanmedian(kpi_14)))





                                print(Y+'hsupa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_15}'+G,
                                    f'Median : {float(np.nanmedian(kpi_15))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_15))) == True:
                                    outSheet.write(
                                    z + 1, 19, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 19, float(np.nanmedian(kpi_15)))

                            


                            except(TypeError):

                                continue

                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'RD3-BH calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 5:

                        
                        outWorkbook = xlsxwriter.Workbook(f"RD4_BH_BL.xlsx")

                        outSheet = outWorkbook.add_worksheet(name='4G_DATA_BH')

                        # ==================== data
                        df_RD4 = pd.read_excel('RD4_BH_data.xlsx', sheet_name='Sheet1')
                        astro_RD4 = df_RD4.to_dict('records')
                        # ====================

                        main_cell_source_index_4 = df_RD4[['cell_ref']].dropna()
                        main_cell_source_index_4 = np.asanyarray(
                            main_cell_source_index_4).flatten()
                        main_cell_source_index_4 = list(
                            np.nan_to_num(main_cell_source_index_4))

                        kpi_list = [
                            
                            'Calculation Period',
                            'Region',
                            'Province',
                            'Cell',
                            'total_traffic_volume(gb)',
                            'Average_Downlink_User_Throughput(Mbit/s)',
                            'Call_Drop_Rate',
                            'Average_UPlink_User_Throughput(Mbit/s)',
                            'RRC_Connection_Setup_Success_Rate_service',
                            'E-RAB_Setup_Success_Rate',
                            'E-RAB_Setup_Success_Rate(Hu_Cell)',
                            'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell',
                            'CSFB_Rate',
                            'S1Signal_E-RAB_Setup_SR(Hu_Cell)',
                            'InterF_HOOut_SR',
                            'IntraF_HOOut_SR',
                            'downlink_cell_throghput(kbit/s)',
                            'uplink_cell_throghput(kbit/s)',
                            'Average_UL_Packet_Loss_%(Huawei_LTE_UCell)',
                            'CSSR(ALL)',
                            'Level'

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

                                if astro_RD4[i]['CELL'] == main_cell_source_index_4[z]:

                                    kpi_1.append(
                                        astro_RD4[i]['Total_Traffic_Volume(GB)'])
                                    kpi_2.append(
                                        astro_RD4[i]['Average_Downlink_User_Throughput(Mbit/s)'])
                                    kpi_3.append(
                                        astro_RD4[i]['Call_Drop_Rate'])
                                    kpi_4.append(astro_RD4[i]['Average_UPlink_User_Throughput(Mbit/s)'])
                                    kpi_5.append(astro_RD4[i]['RRC_Connection_Setup_Success_Rate_service'])
                                    kpi_6.append(astro_RD4[i]['E-RAB_Setup_Success_Rate'])
                                    kpi_7.append(
                                        astro_RD4[i]['E-RAB_Setup_Success_Rate(Hu_Cell)'])
                                    kpi_8.append(
                                        astro_RD4[i]['Intra_RAT_Handover_SR_Intra+Inter_frequency(Huawei_LTE_Cell)'])
                                    kpi_9.append(
                                        astro_RD4[i]['CSFB_Rate'])
                                    kpi_10.append(
                                        astro_RD4[i]['S1Signal_E-RAB_Setup_SR(Hu_Cell)'])
                                    kpi_11.append(astro_RD4[i]['InterF_HOOut_SR'])
                                    kpi_12.append(
                                        astro_RD4[i]['IntraF_HOOut_SR'])
                                    kpi_13.append(
                                        astro_RD4[i]['Downlink_Cell_Throghput(Kbit/s)'])
                                    kpi_14.append(astro_RD4[i]['Uplink_Cell_Throghput(Kbit/s)'])
                                    kpi_15.append(astro_RD4[i]['Average_UL_Packet_Loss_%(Huawei_LTE_UCell)'])
                                    kpi_16.append(astro_RD4[i]['CSSR(ALL)'])
                                

                                else:

                                    continue
                            
                            row_1 = 0

                            row_2 = 0
                            column_2 = 0

                            for kpi in kpi_list:

                                outSheet.write(row_2, column_2, kpi)

                                column_2 += 1



                            try:

                                print(f'{z} cell : {C+main_cell_source_index_4[z]+W}')
                                outSheet.write(
                                    z + 1, 3, main_cell_source_index_4[z])


                                outSheet.write(
                                    z + 1, 0, f"From : {beg_calc} To {end_calc}")



                                outSheet.write(
                                    z + 1, 2, str(main_cell_source_index_4[z][0:2]))

                                    

                                print(Y+'total_traffic_volume(gb)'+W, f'= {kpi_1}'+G,
                                    f'Average : {float(np.nanmean(kpi_1))}'+W)    # edit from here 

                                outSheet.write(z + 1, 4 , float(np.nanmean(kpi_1)))


                                if float(np.nanmean(kpi_1)) > 160:

                                    outSheet.write(                     
                                    z + 1, 20, 'L1')   

                                elif 120 < float(np.nanmean(kpi_1)) <= 160:

                                    outSheet.write(                     
                                    z + 1, 20, 'L2')   
                                
                                elif 80 < float(np.nanmean(kpi_1)) <= 120:

                                    outSheet.write(                     
                                    z + 1, 20, 'L3')   

                                elif 40 < float(np.nanmean(kpi_1)) <= 80:         

                                    outSheet.write(                     
                                    z + 1, 20, 'L4')   

                                elif float(np.nanmean(kpi_1)) <= 40:

                                    outSheet.write(                     
                                    z + 1, 20, 'L5')


                                print(Y+'average_downlink_user_throughput(mbit/s)'+W, f'= {kpi_2}'+G,
                                    f'Median : {float(np.nanmedian(kpi_2))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_2))) == True:
                                    outSheet.write(
                                    z + 1, 5, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 5, float(np.nanmedian(kpi_2)))


                                

                                print(Y+'Call_Drop_Rate'+W, f'= {kpi_3}'+G,
                                    f'Median : {float(np.nanmedian(kpi_3))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_3))) == True:
                                    outSheet.write(
                                    z + 1, 6, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 6, float(np.nanmedian(kpi_3)))





                                print(Y+'Average_UPlink_User_Throughput(Mbit/s)'+W, f'= {kpi_4}'+G,
                                    f'Median : {float(np.nanmedian(kpi_4))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_4))) == True:
                                    outSheet.write(
                                    z + 1, 7, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 7, float(np.nanmedian(kpi_4)))





                                print(Y+'RRC_Connection_Setup_Success_Rate_service'+W, f'= {kpi_5}'+G,
                                    f'Median : {float(np.nanmedian(kpi_5))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_5))) == True:
                                    outSheet.write(
                                    z + 1, 8, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 8, float(np.nanmedian(kpi_5)))




                                print(Y+'E-RAB_Setup_Success_Rate'+W, f'= {kpi_6}'+G,
                                    f'Median : {float(np.nanmedian(kpi_6))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_6))) == True:
                                    outSheet.write(
                                    z + 1, 9, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 9, float(np.nanmedian(kpi_6)))




                                print(Y+'E-RAB_Setup_Success_Rate(Hu_Cell)'+W, f'= {kpi_7}'+G,
                                    f'Median : {float(np.nanmedian(kpi_7))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_7))) == True:
                                    outSheet.write(
                                    z + 1, 10, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 10, float(np.nanmedian(kpi_7)))




                                print(Y+'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell'+W, f'= {kpi_8}'+G,
                                    f'Median : {float(np.nanmedian(kpi_8))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_8))) == True:
                                    outSheet.write(
                                    z + 1, 11, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 11, float(np.nanmedian(kpi_8)))



                                print(Y+'CSFB_Rate'+W, f'= {kpi_9}'+G,
                                    f'Median : {float(np.nanmedian(kpi_9))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_9))) == True:
                                    outSheet.write(
                                    z + 1, 12, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 12, float(np.nanmedian(kpi_9)))




                                print(Y+'S1Signal_E-RAB_Setup_SR(Hu_Cell)'+W, f'= {kpi_10}'+G,
                                    f'Median : {float(np.nanmedian(kpi_10))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_10))) == True:
                                    outSheet.write(
                                    z + 1, 13, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 13, float(np.nanmedian(kpi_10)))



                                print(Y+'InterF_HOOut_SR'+W, f'= {kpi_11}'+G,
                                    f'Median : {float(np.nanmedian(kpi_11))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_11))) == True:
                                    outSheet.write(
                                    z + 1, 14, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 14, float(np.nanmedian(kpi_11)))




                                print(Y+'IntraF_HOOut_SR'+W, f'= {kpi_12}'+G,
                                    f'Median : {float(np.nanmedian(kpi_12))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_12))) == True:
                                    outSheet.write(
                                    z + 1, 15, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 15, float(np.nanmedian(kpi_12)))




                                print(Y+'downlink_cell_throghput(kbit/s)'+W, f'= {kpi_13}'+G,
                                    f'Median : {float(np.nanmedian(kpi_13))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_13))) == True:
                                    outSheet.write(
                                    z + 1, 16, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 16, float(np.nanmedian(kpi_13)))





                                print(Y+'uplink_cell_throghput(kbit/s)'+W, f'= {kpi_14}'+G,
                                    f'Median : {float(np.nanmedian(kpi_14))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_14))) == True:
                                    outSheet.write(
                                    z + 1, 17, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 17, float(np.nanmedian(kpi_14)))




                                print(Y+'Average_UL_Packet_Loss_%(Huawei_LTE_UCell)'+W, f'= {kpi_15}'+G,
                                    f'Median : {float(np.nanmedian(kpi_15))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_15))) == True:
                                    outSheet.write(
                                    z + 1, 18, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 18, float(np.nanmedian(kpi_15)))



                                print(Y+'CSSR(ALL)'+W, f'= {kpi_16}'+G,
                                    f'Median : {float(np.nanmedian(kpi_16))}'+W)

                                if np.isnan(float(np.nanmedian(kpi_16))) == True:
                                    outSheet.write(
                                    z + 1, 19, 'null')        
                                else:
                                    outSheet.write(
                                    z + 1, 19, float(np.nanmedian(kpi_16)))
                                
                                
                                
                            except(TypeError):

                                continue


                        print(R+splitter+W)

                        outWorkbook.close()

                        userfdec = input(
                            C+'RD4-BH calculation Done! continue? [y/n] : '+W)

                        userfdec = userfdec.lower()

                        match userfdec:

                            case 'y':

                                continue

                            case 'n':

                                sys.exit()

                    case 6:

                            os.chdir(gitdir)

                            os.system(fr'python newcell.py')

        
