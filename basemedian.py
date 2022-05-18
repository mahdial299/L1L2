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

user = 'Mehdi Alebrahim'

gitdir = fr'C:\Users\{user}\Desktop\L12'


def ban():

    print(f'''{C+pyfiglet.figlet_format("L1---L2")}
{R+splitter+W}''')


def lister():

    print(G+f'''----------- Quarter calculation ----------

1. 2G voice Calculation
2. 2G data Calculation
3. 3G voice Calculation
4. 3G data Calculation
5. 4G data Calculation

------------------- Add New Cell

6. 

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

                outWorkbook = xlsxwriter.Workbook(f"CC2_Daily_BL.xlsx")

                outSheet = outWorkbook.add_worksheet(name='Median_bl')

                # ================= data
                df_CC2 = pd.read_excel('CC2_data.xlsx', sheet_name='Sheet1')
                astro_CC2 = df_CC2.to_dict('records')
                # ========================

                main_cell_source_index_2 = df_CC2[['cell_ref']].dropna()
                main_cell_source_index_2 = np.asanyarray(
                    main_cell_source_index_2).flatten()
                main_cell_source_index_2 = list(
                    np.nan_to_num(main_cell_source_index_2))

                kpi_list = [

                    'tch_traffic',
                    'available_tch',
                    'htch_traffic',
                    'sdcch_mht',
                    'tch_availability',
                    'amrfr_usage',
                    'amrhr_usage',
                    'cssr3',
                    'sdcch_congestion_rate',
                    'sdcch_drop_rate',
                    'tch_assignment_fr',
                    'tch_cong',
                    'ihsr2',
                    'ohsr2',
                    'sdcch_access_success_rate2',
                    'cdr3',
                    'rx_qualitty_dl_new',
                    'rx_qualitty_ul_new'

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
                            kpi_15.append(
                                astro_CC2[i]['sdcch_access_success_rate2'])
                            kpi_16.append(astro_CC2[i]['cdr3'])
                            kpi_17.append(astro_CC2[i]['rx_qualitty_dl_new'])
                            kpi_18.append(astro_CC2[i]['rx_qualitty_ul_new'])

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

                        print(f'cell : {C+main_cell_source_index_2[z]+W}')

                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_2[z])

                        print(Y+'tch_traffic'+W, f'= {kpi_1}'+G,
                              f'Median : {float(statistics.median(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 1, float(statistics.median(kpi_1)))

                        print(Y+'available_tch'+W, f'= {kpi_2}'+G,
                              f'Median : {float(statistics.median(kpi_2))}'+W)

                        outSheet.write(
                            z + 1, 2, float(statistics.median(kpi_2)))

                        print(Y+'htch_traffic'+W, f'= {kpi_3}'+G,
                              f'Median : {float(statistics.median(kpi_3))}'+W)

                        outSheet.write(
                            z + 1, 3, float(statistics.median(kpi_3)))

                        print(Y+'sdcch_mht'+W, f'= {kpi_4}'+G,
                              f'Median : {float(statistics.median(kpi_4))}'+W)

                        outSheet.write(
                            z + 1, 4, float(statistics.median(kpi_4)))

                        print(Y+'tch_availability'+W, f'= {kpi_5}'+G,
                              f'Median : {float(statistics.median(kpi_5))}'+W)

                        outSheet.write(
                            z + 1, 5, float(statistics.median(kpi_5)))

                        print(Y+'amrfr_usage'+W, f'= {kpi_6}'+G,
                              f'Median : {float(statistics.median(kpi_6))}'+W)

                        outSheet.write(
                            z + 1, 6, float(statistics.median(kpi_6)))

                        print(Y+'amrhr_usage'+W, f'= {kpi_7}'+G,
                              f'Median : {float(statistics.median(kpi_7))}'+W)

                        outSheet.write(
                            z + 1, 7, float(statistics.median(kpi_7)))

                        print(Y+'cssr3'+W, f'= {kpi_8}'+G,
                              f'Median : {float(statistics.median(kpi_8))}'+W)

                        outSheet.write(
                            z + 1, 8, float(statistics.median(kpi_8)))

                        print(Y+'sdcch_congestion_rate'+W, f'= {kpi_9}'+G,
                              f'Median : {float(statistics.median(kpi_9))}'+W)

                        outSheet.write(
                            z + 1, 9, float(statistics.median(kpi_9)))

                        print(Y+'sdcch_drop_rate'+W, f'= {kpi_10}'+G,
                              f'Median : {float(statistics.median(kpi_10))}'+W)

                        outSheet.write(
                            z + 1, 10, float(statistics.median(kpi_10)))

                        print(Y+'tch_assignment_fr'+W, f'= {kpi_11}'+G,
                              f'Median : {float(statistics.median(kpi_11))}'+W)

                        outSheet.write(
                            z + 1, 11, float(statistics.median(kpi_11)))

                        print(Y+'tch_cong'+W, f'= {kpi_12}'+G,
                              f'Median : {float(statistics.median(kpi_12))}'+W)

                        outSheet.write(
                            z + 1, 12, float(statistics.median(kpi_12)))

                        print(Y+'ihsr2'+W, f'= {kpi_13}'+G,
                              f'Median : {float(statistics.median(kpi_13))}'+W)

                        outSheet.write(
                            z + 1, 13, float(statistics.median(kpi_13)))

                        print(Y+'ohsr2'+W, f'= {kpi_14}'+G,
                              f'Median : {float(statistics.median(kpi_14))}'+W)

                        outSheet.write(
                            z + 1, 14, float(statistics.median(kpi_14)))

                        print(Y+'sdcch_access_success_rate2'+W, f'= {kpi_15}'+G,
                              f'Median : {float(statistics.median(kpi_15))}'+W)

                        outSheet.write(
                            z + 1, 15, float(statistics.median(kpi_15)))

                        print(Y+'cdr3'+W, f'= {kpi_16}'+G,
                              f'Median : {float(statistics.median(kpi_16))}'+W)

                        outSheet.write(
                            z + 1, 16, float(statistics.median(kpi_16)))

                        print(Y+'rx_qualitty_dl_new'+W, f'= {kpi_17}'+G,
                              f'Median : {float(statistics.median(kpi_17))}'+W)

                        outSheet.write(
                            z + 1, 17, float(statistics.median(kpi_17)))

                        print(Y+'rx_qualitty_ul_new'+W, f'= {kpi_18}'+G,
                              f'Median : {float(statistics.median(kpi_18))}'+W)

                        outSheet.write(
                            z + 1, 18, float(statistics.median(kpi_18)))

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
                outWorkbook = xlsxwriter.Workbook(f"RD2_BL.xlsx")

                outSheet = outWorkbook.add_worksheet(name='Median_bl')

                # ================= data
                df_RD2 = pd.read_excel('RD2_data.xlsx', sheet_name='Sheet1')
                astro_RD2 = df_RD2.to_dict('records')
                # ========================

                main_cell_source_index_2 = df_RD2[['cell_ref']].dropna()
                main_cell_source_index_2 = np.asanyarray(
                    main_cell_source_index_2).flatten()
                main_cell_source_index_2 = list(
                    np.nan_to_num(main_cell_source_index_2))

                kpi_list = [
                    'tbf_establishment_success_rate(ul+dl)(%)(hu_cell)',
                    'tbf_drop(ul+dl)(hu_cell)',
                    'average_throughput_of_downlink_gprs_llc_per_user(kbps)',
                    'average_throughput_of_downlink_egprs_llc_per_user(kbps)',
                    'thr_dl_gprs_per_ts(cell_hu)',
                    'thr_dl_egprs_per_ts(cell_hu)',
                    'payload_total_ul(cell_hu)',
                    'payload_total_dl(cell_hu)',
                    'payload_total(cell_hu)',
                    'edge_share_payload(cell_hu)',
                    'tch_availability(hu_cell)',
                    'trx'
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
                    kpi_9 = []
                    kpi_10 = []
                    kpi_11 = []
                    kpi_12 = []

                    for i in range(len(astro_RD2)):

                        if astro_RD2[i]['cell'] == main_cell_source_index_2[z]:

                            kpi_1.append(
                                astro_RD2[i]['tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'])
                            kpi_2.append(
                                astro_RD2[i]['tbf_drop(ul+dl)(hu_cell)'])
                            kpi_3.append(
                                astro_RD2[i]['average_throughput_of_downlink_gprs_llc_per_user(kbps)'])
                            kpi_4.append(
                                astro_RD2[i]['average_throughput_of_downlink_egprs_llc_per_user(kbps)'])
                            kpi_5.append(
                                astro_RD2[i]['thr_dl_gprs_per_ts(cell_hu)'])
                            kpi_6.append(
                                astro_RD2[i]['thr_dl_egprs_per_ts(cell_hu)'])
                            kpi_7.append(
                                astro_RD2[i]['payload_total_ul(cell_hu)'])
                            kpi_8.append(
                                astro_RD2[i]['payload_total_dl(cell_hu)'])
                            kpi_9.append(
                                astro_RD2[i]['payload_total(cell_hu)'])
                            kpi_10.append(
                                astro_RD2[i]['edge_share_payload(cell_hu)'])
                            kpi_11.append(
                                astro_RD2[i]['tch_availability(hu_cell)'])
                            kpi_12.append(astro_RD2[i]['trx'])

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

                        print(f'cell : {C+main_cell_source_index_2[z]+W}')
                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_2[z])

                        print(Y+'tbf_establishment_success_rate(ul+dl)(%)(hu_cell)'+W, f'= {kpi_1}'+G,
                              f'Median : {float(statistics.median(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 1, float(statistics.median(kpi_1)))

                        print(Y+'tbf_drop(ul+dl)(hu_cell)'+W, f'= {kpi_2}'+G,
                              f'Median : {float(statistics.median(kpi_2))}'+W)

                        outSheet.write(
                            z + 1, 2, float(statistics.median(kpi_2)))

                        print(Y+'average_throughput_of_downlink_gprs_llc_per_user(kbps)'+W, f'= {kpi_3}'+G,
                              f'Median : {float(statistics.median(kpi_3))}'+W)

                        outSheet.write(
                            z + 1, 3, float(statistics.median(kpi_3)))

                        print(Y+'average_throughput_of_downlink_egprs_llc_per_user(kbps)'+W, f'= {kpi_4}'+G,
                              f'Median : {float(statistics.median(kpi_4))}'+W)

                        outSheet.write(
                            z + 1, 4, float(statistics.median(kpi_4)))

                        print(Y+'thr_dl_gprs_per_ts(cell_hu)'+W, f'= {kpi_5}'+G,
                              f'Median : {float(statistics.median(kpi_5))}'+W)

                        outSheet.write(
                            z + 1, 5, float(statistics.median(kpi_5)))

                        print(Y+'thr_dl_egprs_per_ts(cell_hu)'+W, f'= {kpi_6}'+G,
                              f'Median : {float(statistics.median(kpi_6))}'+W)

                        outSheet.write(
                            z + 1, 6, float(statistics.median(kpi_6)))

                        print(Y+'payload_total_ul(cell_hu)'+W, f'= {kpi_7}'+G,
                              f'Median : {float(statistics.median(kpi_7))}'+W)

                        outSheet.write(
                            z + 1, 7, float(statistics.median(kpi_7)))

                        print(Y+'payload_total_dl(cell_hu)'+W, f'= {kpi_8}'+G,
                              f'Median : {float(statistics.median(kpi_8))}'+W)

                        outSheet.write(
                            z + 1, 8, float(statistics.median(kpi_8)))

                        print(Y+'payload_total(cell_hu)'+W, f'= {kpi_9}'+G,
                              f'Median : {float(statistics.median(kpi_9))}'+W)

                        outSheet.write(
                            z + 1, 9, float(statistics.median(kpi_9)))

                        print(Y+'edge_share_payload(cell_hu)'+W, f'= {kpi_10}'+G,
                              f'Median : {float(statistics.median(kpi_10))}'+W)

                        outSheet.write(
                            z + 1, 10, float(statistics.median(kpi_10)))

                        print(Y+'tch_availability(hu_cell)'+W, f'= {kpi_11}'+G,
                              f'Median : {float(statistics.median(kpi_11))}'+W)

                        outSheet.write(
                            z + 1, 11, float(statistics.median(kpi_11)))

                        print(Y+'trx'+W, f'= {kpi_12}'+G,
                              f'Median : {float(statistics.median(kpi_12))}'+W)

                        outSheet.write(
                            z + 1, 12, float(statistics.median(kpi_12)))

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
                df_CC3 = pd.read_excel('CC3_data.xlsx', sheet_name='Sheet1')
                astro_CC3 = df_CC3.to_dict('records')
                # ====================

                main_cell_source_index_3 = df_CC3[['cell_ref']].dropna()
                main_cell_source_index_3 = np.asanyarray(
                    main_cell_source_index_3).flatten()
                main_cell_source_index_3 = list(
                    np.nan_to_num(main_cell_source_index_3))

                kpi_list = [

                    'cs_erlang',
                    'cs_rrc_connection_establishment_sr',
                    'cs_rab_setup_success_ratio',
                    'softer_handover_success_ratio(hu_cell)',
                    'cs_rab_setup_congestion_rate(hu_cell)',
                    'radio_network_availability_ratio(hu_cell)',
                    'bler_amr(cell_huawei)',
                    'cs_irat_ho_sr',
                    'amr_call_drop_ratio_new(hu_cell)',
                    'csps_rab_setup_success_ratio',
                    'interfrequency_hardhandover_success_ratio_csservice',
                    'cs_cssr',
                    'rrc_setup_success_ratio(cell.service)',
                    'soft_handover_succ_rate',
                    'inter_carrier_ho_success_rate',
                    'cs_rrc_setup_sr_ura_pch(hu_cell)',
                    'cs_cssr_ura_pch(hu_cell)'

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
                            kpi_2.append(
                                astro_CC3[i]['cs_rrc_connection_establishment_sr'])
                            kpi_3.append(
                                astro_CC3[i]['cs_rab_setup_success_ratio'])
                            kpi_4.append(
                                astro_CC3[i]['softer_handover_success_ratio(hu_cell)'])
                            kpi_5.append(
                                astro_CC3[i]['cs_rab_setup_congestion_rate(hu_cell)'])
                            kpi_6.append(
                                astro_CC3[i]['radio_network_availability_ratio(hu_cell)'])
                            kpi_7.append(astro_CC3[i]['bler_amr(cell_huawei)'])
                            kpi_8.append(astro_CC3[i]['cs_irat_ho_sr'])
                            kpi_9.append(
                                astro_CC3[i]['amr_call_drop_ratio_new(hu_cell)'])
                            kpi_10.append(
                                astro_CC3[i]['csps_rab_setup_success_ratio'])
                            kpi_11.append(
                                astro_CC3[i]['interfrequency_hardhandover_success_ratio_csservice'])
                            kpi_12.append(astro_CC3[i]['cs_cssr'])
                            kpi_13.append(
                                astro_CC3[i]['rrc_setup_success_ratio(cell.service)'])
                            kpi_14.append(
                                astro_CC3[i]['soft_handover_succ_rate'])
                            kpi_15.append(
                                astro_CC3[i]['inter_carrier_ho_success_rate'])
                            kpi_16.append(
                                astro_CC3[i]['cs_rrc_setup_sr_ura_pch(hu_cell)'])
                            kpi_17.append(
                                astro_CC3[i]['cs_cssr_ura_pch(hu_cell)'])

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

                        print(f'cell : {C+main_cell_source_index_3[z]+W}')
                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_3[z])

                        print(Y+'cs_erlang'+W, f'= {kpi_1}'+G,
                              f'Median : {float(statistics.median(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 1, float(statistics.median(kpi_1)))

                        print(Y+'cs_rrc_connection_establishment_sr'+W, f'= {kpi_2}'+G,
                              f'Median : {float(statistics.median(kpi_2))}'+W)

                        outSheet.write(
                            z + 1, 2, float(statistics.median(kpi_2)))

                        print(Y+'cs_rab_setup_success_ratio'+W, f'= {kpi_3}'+G,
                              f'Median : {float(statistics.median(kpi_3))}'+W)

                        outSheet.write(
                            z + 1, 3, float(statistics.median(kpi_3)))

                        print(Y+'softer_handover_success_ratio(hu_cell)'+W, f'= {kpi_4}'+G,
                              f'Median : {float(statistics.median(kpi_4))}'+W)

                        outSheet.write(
                            z + 1, 4, float(statistics.median(kpi_4)))

                        print(Y+'cs_rab_setup_congestion_rate(hu_cell)'+W, f'= {kpi_5}'+G,
                              f'Median : {float(statistics.median(kpi_5))}'+W)

                        outSheet.write(
                            z + 1, 5, float(statistics.median(kpi_5)))

                        print(Y+'radio_network_availability_ratio(hu_cell)'+W, f'= {kpi_6}'+G,
                              f'Median : {float(statistics.median(kpi_6))}'+W)

                        outSheet.write(
                            z + 1, 6, float(statistics.median(kpi_6)))

                        print(Y+'bler_amr(cell_huawei)'+W, f'= {kpi_7}'+G,
                              f'Median : {float(statistics.median(kpi_7))}'+W)

                        outSheet.write(
                            z + 1, 7, float(statistics.median(kpi_7)))

                        print(Y+'cs_irat_ho_sr'+W, f'= {kpi_8}'+G,
                              f'Median : {float(statistics.median(kpi_8))}'+W)

                        outSheet.write(
                            z + 1, 8, float(statistics.median(kpi_8)))

                        print(Y+'amr_call_drop_ratio_new(hu_cell)'+W, f'= {kpi_9}'+G,
                              f'Median : {float(statistics.median(kpi_9))}'+W)

                        outSheet.write(
                            z + 1, 9, float(statistics.median(kpi_9)))

                        print(Y+'csps_rab_setup_success_ratio'+W, f'= {kpi_10}'+G,
                              f'Median : {float(statistics.median(kpi_10))}'+W)

                        outSheet.write(
                            z + 1, 10, float(statistics.median(kpi_10)))

                        print(Y+'interfrequency_hardhandover_success_ratio_csservice'+W, f'= {kpi_11}'+G,
                              f'Median : {float(statistics.median(kpi_11))}'+W)

                        outSheet.write(
                            z + 1, 11, float(statistics.median(kpi_11)))

                        print(Y+'cs_cssr'+W, f'= {kpi_12}'+G,
                              f'Median : {float(statistics.median(kpi_12))}'+W)

                        outSheet.write(
                            z + 1, 12, float(statistics.median(kpi_12)))

                        print(Y+'rrc_setup_success_ratio(cell.service)'+W, f'= {kpi_13}'+G,
                              f'Median : {float(statistics.median(kpi_13))}'+W)

                        outSheet.write(
                            z + 1, 13, float(statistics.median(kpi_13)))

                        print(Y+'soft_handover_succ_rate'+W, f'= {kpi_14}'+G,
                              f'Median : {float(statistics.median(kpi_14))}'+W)

                        outSheet.write(
                            z + 1, 14, float(statistics.median(kpi_14)))

                        print(Y+'inter_carrier_ho_success_rate'+W, f'= {kpi_15}'+G,
                              f'Median : {float(statistics.median(kpi_15))}'+W)

                        outSheet.write(
                            z + 1, 15, float(statistics.median(kpi_15)))

                        print(Y+'cs_rrc_setup_sr_ura_pch(hu_cell)'+W, f'= {kpi_16}'+G,
                              f'Median : {float(statistics.median(kpi_16))}'+W)

                        outSheet.write(
                            z + 1, 16, float(statistics.median(kpi_16)))

                        print(Y+'cs_cssr_ura_pch(hu_cell)'+W, f'= {kpi_17}'+G,
                              f'Median : {float(statistics.median(kpi_17))}'+W)

                        outSheet.write(
                            z + 1, 17, float(statistics.median(kpi_17)))

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
                df_RD3 = pd.read_excel('RD3_data.xlsx', sheet_name='Sheet1')
                astro_RD3 = df_RD3.to_dict('records')
                # ====================

                main_cell_source_index_3 = df_RD3[['cell_ref']].dropna()
                main_cell_source_index_3 = np.asanyarray(
                    main_cell_source_index_3).flatten()
                main_cell_source_index_3 = list(
                    np.nan_to_num(main_cell_source_index_3))

                kpi_list = [

                    'payload',
                    'ps_cssr',
                    'ps_call_drop_ratio',
                    'average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)',
                    'hsupa_uplink_throughput_in_v16(cell_hu)',
                    'cs+ps_rab_setup_success_ratio',
                    'hsdpa_soft_handover_success_ratio',
                    'hs_share_payload_%',
                    'hsdpa_cdr(%)_(hu_cell)_new',
                    'hsupa_cdr(%)_(hu_cell)_new',
                    'ps_r99_call_drop_ratio_with_pch(hu_cell)',
                    'nack_ratio(cell_huawei)',
                    'hsdpa_scheduling_cell_throughput(cell_huawei)',
                    'hsupa_cell_throughput(kbps)(hu_cell)',
                    'radio_network_availability_ratio(hu_cell)',
                    'ps_rab_setup_success_ratio(hu_cell)',
                    'bler9',
                    'cqi>20',
                    'ps_rrc_connection_success_rate_repeatless(hu_cell)',
                    'ps_r99_rab_setup_success_ratio(hu_cell)',
                    'hsdpa_rab_setup_success_ratio(hu_cell)',
                    'hsupa_rab_setup_success_ratio(hu_cell)',
                    'vs.rab.abnormrel.ps_rnc',
                    'ps_rab_setup_congestion_rate',
                    'ps_rab_setup_success_ratio',
                    'ps_rab_congestion_rate',
                    'hsdpa_user_throughput',
                    'hsupa_throughput_mace',
                    'ps_cssr_ura_pch(hu_cell)',
                    'pch2dch_statetrans_sr(hu_cell)',
                    'mean_rtwp(cell_hu)',
                    'cqi_new(hu_cell)'

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
                                astro_RD3[i]['hsupa_cdr(%)_(hu_cell)_new'])
                            kpi_11.append(
                                astro_RD3[i]['ps_r99_call_drop_ratio_with_pch(hu_cell)'])
                            kpi_12.append(
                                astro_RD3[i]['nack_ratio(cell_huawei)'])
                            kpi_13.append(
                                astro_RD3[i]['hsdpa_scheduling_cell_throughput(cell_huawei)'])
                            kpi_14.append(
                                astro_RD3[i]['hsupa_cell_throughput(kbps)(hu_cell)'])
                            kpi_15.append(
                                astro_RD3[i]['radio_network_availability_ratio(hu_cell)'])
                            kpi_16.append(
                                astro_RD3[i]['ps_rab_setup_success_ratio(hu_cell)'])
                            kpi_17.append(astro_RD3[i]['bler9'])
                            kpi_18.append(astro_RD3[i]['cqi>20'])
                            kpi_19.append(
                                astro_RD3[i]['ps_rrc_connection_success_rate_repeatless(hu_cell)'])
                            kpi_20.append(
                                astro_RD3[i]['ps_r99_rab_setup_success_ratio(hu_cell)'])
                            kpi_21.append(
                                astro_RD3[i]['hsdpa_rab_setup_success_ratio(hu_cell)'])
                            kpi_22.append(
                                astro_RD3[i]['hsupa_rab_setup_success_ratio(hu_cell)'])
                            kpi_23.append(
                                astro_RD3[i]['vs.rab.abnormrel.ps_rnc'])
                            kpi_24.append(
                                astro_RD3[i]['ps_rab_setup_congestion_rate'])
                            kpi_25.append(
                                astro_RD3[i]['ps_rab_setup_success_ratio'])
                            kpi_26.append(
                                astro_RD3[i]['ps_rab_congestion_rate'])
                            kpi_27.append(
                                astro_RD3[i]['hsdpa_user_throughput'])
                            kpi_28.append(
                                astro_RD3[i]['hsupa_throughput_mace'])
                            kpi_29.append(
                                astro_RD3[i]['ps_cssr_ura_pch(hu_cell)'])
                            kpi_30.append(
                                astro_RD3[i]['pch2dch_statetrans_sr(hu_cell)'])
                            kpi_31.append(astro_RD3[i]['mean_rtwp(cell_hu)'])
                            kpi_32.append(astro_RD3[i]['cqi_new(hu_cell)'])

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

                        print(f'cell : {C+main_cell_source_index_3[z]+W}')
                        outSheet.write(
                            z + 1, row_1, main_cell_source_index_3[z])

                        print(Y+'payload'+W, f'= {kpi_1}'+G,
                              f'Median : {float(statistics.median(kpi_1))}'+W)

                        outSheet.write(
                            z + 1, 1, float(statistics.median(kpi_1)))

                        print(Y+'ps_cssr'+W, f'= {kpi_2}'+G,
                              f'Median : {float(statistics.median(kpi_2))}'+W)

                        outSheet.write(
                            z + 1, 2, float(statistics.median(kpi_2)))

                        print(Y+'ps_call_drop_ratio'+W, f'= {kpi_3}'+G,
                              f'Median : {float(statistics.median(kpi_3))}'+W)

                        outSheet.write(
                            z + 1, 3, float(statistics.median(kpi_3)))

                        print(Y+'average_hsdpa_user_throughput_dc+sc(mbit/s)(cell_huawei)'+W, f'= {kpi_4}'+G,
                              f'Median : {float(statistics.median(kpi_4))}'+W)

                        outSheet.write(
                            z + 1, 4, float(statistics.median(kpi_4)))

                        print(Y+'hsupa_uplink_throughput_in_v16(cell_hu)'+W, f'= {kpi_5}'+G,
                              f'Median : {float(statistics.median(kpi_5))}'+W)

                        outSheet.write(
                            z + 1, 5, float(statistics.median(kpi_5)))

                        print(Y+'cs+ps_rab_setup_success_ratio'+W, f'= {kpi_6}'+G,
                              f'Median : {float(statistics.median(kpi_6))}'+W)

                        outSheet.write(
                            z + 1, 6, float(statistics.median(kpi_6)))

                        print(Y+'hsdpa_soft_handover_success_ratio'+W, f'= {kpi_7}'+G,
                              f'Median : {float(statistics.median(kpi_7))}'+W)

                        outSheet.write(
                            z + 1, 7, float(statistics.median(kpi_7)))

                        print(Y+'hs_share_payload_%'+W, f'= {kpi_8}'+G,
                              f'Median : {float(statistics.median(kpi_8))}'+W)

                        outSheet.write(
                            z + 1, 8, float(statistics.median(kpi_8)))

                        print(Y+'hsdpa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_9}'+G,
                              f'Median : {float(statistics.median(kpi_9))}'+W)

                        outSheet.write(
                            z + 1, 9, float(statistics.median(kpi_9)))

                        print(Y+'hsupa_cdr(%)_(hu_cell)_new'+W, f'= {kpi_10}'+G,
                              f'Median : {float(statistics.median(kpi_10))}'+W)

                        outSheet.write(
                            z + 1, 10, float(statistics.median(kpi_10)))

                        print(Y+'ps_r99_call_drop_ratio_with_pch(hu_cell)'+W, f'= {kpi_11}'+G,
                              f'Median : {float(statistics.median(kpi_11))}'+W)

                        outSheet.write(
                            z + 1, 11, float(statistics.median(kpi_11)))

                        print(Y+'nack_ratio(cell_huawei)'+W, f'= {kpi_12}'+G,
                              f'Median : {float(statistics.median(kpi_12))}'+W)

                        outSheet.write(
                            z + 1, 12, float(statistics.median(kpi_12)))

                        print(Y+'hsdpa_scheduling_cell_throughput(cell_huawei)'+W, f'= {kpi_13}'+G,
                              f'Median : {float(statistics.median(kpi_13))}'+W)

                        outSheet.write(
                            z + 1, 13, float(statistics.median(kpi_13)))

                        print(Y+'hsupa_cell_throughput(kbps)(hu_cell)'+W, f'= {kpi_14}'+G,
                              f'Median : {float(statistics.median(kpi_14))}'+W)

                        outSheet.write(
                            z + 1, 14, float(statistics.median(kpi_14)))

                        print(Y+'radio_network_availability_ratio(hu_cell)'+W, f'= {kpi_15}'+G,
                              f'Median : {float(statistics.median(kpi_15))}'+W)

                        outSheet.write(
                            z + 1, 15, float(statistics.median(kpi_15)))

                        print(Y+'ps_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_16}'+G,
                              f'Median : {float(statistics.median(kpi_16))}'+W)

                        outSheet.write(
                            z + 1, 16, float(statistics.median(kpi_16)))

                        print(Y+'bler9'+W, f'= {kpi_17}'+G,
                              f'Median : {float(statistics.median(kpi_17))}'+W)

                        outSheet.write(
                            z + 1, 17, float(statistics.median(kpi_17)))

                        print(Y+'cqi>20'+W, f'= {kpi_18}'+G,
                              f'Median : {float(statistics.median(kpi_18))}'+W)

                        outSheet.write(
                            z + 1, 18, float(statistics.median(kpi_18)))

                        print(Y+'ps_rrc_connection_success_rate_repeatless(hu_cell)'+W, f'= {kpi_19}'+G,
                              f'Median : {float(statistics.median(kpi_19))}'+W)

                        outSheet.write(
                            z + 1, 19, float(statistics.median(kpi_19)))

                        print(Y+'ps_r99_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_20}'+G,
                              f'Median : {float(statistics.median(kpi_20))}'+W)

                        outSheet.write(
                            z + 1, 20, float(statistics.median(kpi_20)))

                        print(Y+'hsdpa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_21}'+G,
                              f'Median : {float(statistics.median(kpi_21))}'+W)

                        outSheet.write(
                            z + 1, 21, float(statistics.median(kpi_21)))

                        print(Y+'hsupa_rab_setup_success_ratio(hu_cell)'+W, f'= {kpi_22}'+G,
                              f'Median : {float(statistics.median(kpi_22))}'+W)

                        outSheet.write(
                            z + 1, 22, float(statistics.median(kpi_22)))

                        print(Y+'vs.rab.abnormrel.ps_rnc'+W, f'= {kpi_23}'+G,
                              f'Median : {float(statistics.median(kpi_23))}'+W)

                        outSheet.write(
                            z + 1, 23, float(statistics.median(kpi_23)))

                        print(Y+'ps_rab_setup_congestion_rate'+W, f'= {kpi_24}'+G,
                              f'Median : {float(statistics.median(kpi_24))}'+W)

                        outSheet.write(
                            z + 1, 24, float(statistics.median(kpi_24)))

                        print(Y+'ps_rab_setup_success_ratio'+W, f'= {kpi_25}'+G,
                              f'Median : {float(statistics.median(kpi_25))}'+W)

                        outSheet.write(
                            z + 1, 25, float(statistics.median(kpi_25)))

                        print(Y+'ps_rab_congestion_rate'+W, f'= {kpi_26}'+G,
                              f'Median : {float(statistics.median(kpi_26))}'+W)

                        outSheet.write(
                            z + 1, 26, float(statistics.median(kpi_26)))

                        print(Y+'hsdpa_user_throughput'+W, f'= {kpi_27}'+G,
                              f'Median : {float(statistics.median(kpi_27))}'+W)

                        outSheet.write(
                            z + 1, 27, float(statistics.median(kpi_27)))

                        print(Y+'hsupa_throughput_mace'+W, f'= {kpi_28}'+G,
                              f'Median : {float(statistics.median(kpi_28))}'+W)

                        outSheet.write(
                            z + 1, 28, float(statistics.median(kpi_28)))

                        print(Y+'ps_cssr_ura_pch(hu_cell)'+W, f'= {kpi_29}'+G,
                              f'Median : {float(statistics.median(kpi_29))}'+W)

                        outSheet.write(
                            z + 1, 29, float(statistics.median(kpi_29)))

                        print(Y+'pch2dch_statetrans_sr(hu_cell)'+W, f'= {kpi_30}'+G,
                              f'Median : {float(statistics.median(kpi_30))}'+W)

                        outSheet.write(
                            z + 1, 30, float(statistics.median(kpi_30)))

                        print(Y+'mean_rtwp(cell_hu)'+W, f'= {kpi_31}'+G,
                              f'Median : {float(statistics.median(kpi_31))}'+W)

                        outSheet.write(
                            z + 1, 31, float(statistics.median(kpi_31)))

                        print(Y+'cqi_new(hu_cell)'+W, f'= {kpi_32}'+G,
                              f'Median : {float(statistics.median(kpi_32))}'+W)

                        outSheet.write(
                            z + 1, 32, float(statistics.median(kpi_32)))

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
                df_RD4 = pd.read_excel('RD4_data.xlsx', sheet_name='Sheet1')
                astro_RD4 = df_RD4.to_dict('records')
                # =====================

                main_cell_source_index_4 = df_RD4[['cell_ref']].dropna()
                main_cell_source_index_4 = np.asanyarray(
                    main_cell_source_index_4).flatten()
                main_cell_source_index_4 = list(
                    np.nan_to_num(main_cell_source_index_4))

                kpi_list = [

                    'total_traffic_volume(gb)',
                    'e-rab_setup_success_rate(hu_cell)',
                    'e-rab_setup_success_rate',
                    'ran_avail_rate',
                    'interf_hoout_sr',
                    'intraf_hoout_sr',
                    'inter_rat_handover_out_success_rate(3gpltogsm)',
                    'inter_rat_handover_out_successrate(3gpltowcdma)',
                    'average_dl_latency_ms(huawei_lte_eucell)',
                    'average_ul_packet_loss_%(huawei_lte_ucell)',
                    'call_drop_rate',
                    'average_downlink_user_throughput(mbit/s)',
                    'average_uplink_user_throughput(mbit/s)',
                    'csfb_rate',
                    'cssr(all)',
                    'downlink_traffic_volume(gb)',
                    'ul_traffic_volume(gb)',
                    'downlink_cell_throghput(kbit/s)',
                    'uplink_cell_throghput(kbit/s)',
                    'max_no_user',
                    'number_of_available_downlink_prbs_cell',
                    'average_cqi(huawei_lte_cell)',
                    'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell',
                    'rrc_connection_setup_success_rate_service',
                    's1signal_e-rab_setup_sr(hu_cell)',
                    'cell_unvail_duration_daily(huawei_cell_lte)',
                    'rssi_pucch(huawei_lte_cell)',
                    'rssi_pusch(huawei_lte_cell)',
                    'cell_availability_rate_exclude_blocking(cell_hu)',
                    'cell_availability_rate_include_blocking(cell_hu)',
                    'cell_availability_rate_include_blocking',
                    'cell_availability_rate_include_blocking(cell_hu_no_null)'

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

                    for i in range(len(astro_RD4)):

                        if astro_RD4[i]['cell'] == main_cell_source_index_4[z]:

                            kpi_1.append(
                                astro_RD4[i]['total_traffic_volume(gb)'])
                            kpi_2.append(
                                astro_RD4[i]['e-rab_setup_success_rate(hu_cell)'])
                            kpi_3.append(
                                astro_RD4[i]['e-rab_setup_success_rate'])
                            kpi_4.append(astro_RD4[i]['ran_avail_rate'])
                            kpi_5.append(astro_RD4[i]['interf_hoout_sr'])
                            kpi_6.append(astro_RD4[i]['intraf_hoout_sr'])
                            kpi_7.append(
                                astro_RD4[i]['inter_rat_handover_out_success_rate(3gpltogsm)'])
                            kpi_8.append(
                                astro_RD4[i]['inter_rat_handover_out_successrate(3gpltowcdma)'])
                            kpi_9.append(
                                astro_RD4[i]['average_dl_latency_ms(huawei_lte_eucell)'])
                            kpi_10.append(
                                astro_RD4[i]['average_ul_packet_loss_%(huawei_lte_ucell)'])
                            kpi_11.append(astro_RD4[i]['call_drop_rate'])
                            kpi_12.append(
                                astro_RD4[i]['average_downlink_user_throughput(mbit/s)'])
                            kpi_13.append(
                                astro_RD4[i]['average_uplink_user_throughput(mbit/s)'])
                            kpi_14.append(astro_RD4[i]['csfb_rate'])
                            kpi_15.append(astro_RD4[i]['cssr(all)'])
                            kpi_16.append(
                                astro_RD4[i]['downlink_traffic_volume(gb)'])
                            kpi_17.append(
                                astro_RD4[i]['ul_traffic_volume(gb)'])
                            kpi_18.append(
                                astro_RD4[i]['downlink_cell_throghput(kbit/s)'])
                            kpi_19.append(
                                astro_RD4[i]['uplink_cell_throghput(kbit/s)'])
                            kpi_20.append(astro_RD4[i]['max_no_user'])
                            kpi_21.append(
                                astro_RD4[i]['number_of_available_downlink_prbs_cell'])
                            kpi_22.append(
                                astro_RD4[i]['average_cqi(huawei_lte_cell)'])
                            kpi_23.append(
                                astro_RD4[i]['intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell'])
                            kpi_24.append(
                                astro_RD4[i]['rrc_connection_setup_success_rate_service'])
                            kpi_25.append(
                                astro_RD4[i]['s1signal_e-rab_setup_sr(hu_cell)'])
                            kpi_26.append(
                                astro_RD4[i]['cell_unvail_duration_daily(huawei_cell_lte)'])
                            kpi_27.append(
                                astro_RD4[i]['rssi_pucch(huawei_lte_cell)'])
                            kpi_28.append(
                                astro_RD4[i]['rssi_pusch(huawei_lte_cell)'])
                            kpi_29.append(
                                astro_RD4[i]['cell_availability_rate_exclude_blocking(cell_hu)'])
                            kpi_30.append(
                                astro_RD4[i]['cell_availability_rate_include_blocking(cell_hu)'])
                            kpi_31.append(
                                astro_RD4[i]['cell_availability_rate_include_blocking'])
                            kpi_32.append(
                                astro_RD4[i]['cell_availability_rate_include_blocking(cell_hu_no_null)'])

                        else:

                            continue
                    
                    row_1 = 0

                    row_2 = 0
                    column_2 = 1

                    for kpi in kpi_list:

                        outSheet.write(row_2, column_2, kpi)

                        column_2 += 1

                    try:

                        outSheet.write(row_2, row_1, 'cell')
                        print(f'cell : {C+main_cell_source_index_4[z]+W}')
                        outSheet.write(z + 1, row_1 , main_cell_source_index_4[z])

                        print(Y+'total_traffic_volume(gb)'+W, f'= {kpi_1}'+G,
                            f'Median : {float(statistics.median(kpi_1))}'+W)

                        outSheet.write(z + 1, 1 , float(statistics.median(kpi_1)))

                        print(Y+'e-rab_setup_success_rate(hu_cell)'+W, f'= {kpi_2}'+G,
                            f'Median : {float(statistics.median(kpi_2))}'+W)

                        outSheet.write(z + 1, 2 , float(statistics.median(kpi_2)))

                        print(Y+'e-rab_setup_success_rate'+W, f'= {kpi_3}'+G,
                            f'Median : {float(statistics.median(kpi_3))}'+W)

                        outSheet.write(z + 1, 3 , float(statistics.median(kpi_3)))

                        print(Y+'ran_avail_rate'+W, f'= {kpi_4}'+G,
                            f'Median : {float(statistics.median(kpi_4))}'+W)

                        outSheet.write(z + 1, 4 , float(statistics.median(kpi_4)))

                        print(Y+'interf_hoout_sr'+W, f'= {kpi_5}'+G,
                            f'Median : {float(statistics.median(kpi_5))}'+W)

                        outSheet.write(z + 1, 5 , float(statistics.median(kpi_5)))

                        print(Y+'intraf_hoout_sr'+W, f'= {kpi_6}'+G,
                            f'Median : {float(statistics.median(kpi_6))}'+W)

                        outSheet.write(z + 1, 6 , float(statistics.median(kpi_6)))

                        print(Y+'inter_rat_handover_out_success_rate(3gpltogsm)'+W, f'= {kpi_7}'+G,
                            f'Median : {float(statistics.median(kpi_7))}'+W)

                        outSheet.write(z + 1, 7 , float(statistics.median(kpi_7)))

                        print(Y+'inter_rat_handover_out_successrate(3gpltowcdma)'+W, f'= {kpi_8}'+G,
                            f'Median : {float(statistics.median(kpi_8))}'+W)

                        outSheet.write(z + 1, 8 , float(statistics.median(kpi_8)))

                        print(Y+'average_dl_latency_ms(huawei_lte_eucell)'+W, f'= {kpi_9}'+G,
                            f'Median : {float(statistics.median(kpi_9))}'+W)

                        outSheet.write(z + 1, 9 , float(statistics.median(kpi_9)))

                        print(Y+'average_ul_packet_loss_%(huawei_lte_ucell)'+W, f'= {kpi_10}'+G,
                            f'Median : {float(statistics.median(kpi_10))}'+W)

                        outSheet.write(z + 1, 10 , float(statistics.median(kpi_10)))

                        print(Y+'call_drop_rate'+W, f'= {kpi_11}'+G,
                            f'Median : {float(statistics.median(kpi_11))}'+W)

                        outSheet.write(z + 1, 11 , float(statistics.median(kpi_11)))

                        print(Y+'average_downlink_user_throughput(mbit/s)'+W, f'= {kpi_12}'+G,
                            f'Median : {float(statistics.median(kpi_12))}'+W)

                        outSheet.write(z + 1, 12 , float(statistics.median(kpi_12)))

                        print(Y+'average_uplink_user_throughput(mbit/s)'+W, f'= {kpi_13}'+G,
                            f'Median : {float(statistics.median(kpi_13))}'+W)

                        outSheet.write(z + 1, 13 , float(statistics.median(kpi_13)))

                        print(Y+'csfb_rate'+W, f'= {kpi_14}'+G,
                            f'Median : {float(statistics.median(kpi_14))}'+W)

                        outSheet.write(z + 1, 14 , float(statistics.median(kpi_14)))

                        print(Y+'cssr(all)'+W, f'= {kpi_15}'+G,
                            f'Median : {float(statistics.median(kpi_15))}'+W)

                        outSheet.write(z + 1, 15 , float(statistics.median(kpi_15)))

                        print(Y+'downlink_traffic_volume(gb)'+W, f'= {kpi_16}'+G,
                            f'Median : {float(statistics.median(kpi_16))}'+W)

                        outSheet.write(z + 1, 16 , float(statistics.median(kpi_16)))

                        print(Y+'ul_traffic_volume(gb)'+W, f'= {kpi_17}'+G,
                            f'Median : {float(statistics.median(kpi_17))}'+W)

                        outSheet.write(z + 1, 17 , float(statistics.median(kpi_17)))

                        print(Y+'downlink_cell_throghput(kbit/s)'+W, f'= {kpi_18}'+G,
                            f'Median : {float(statistics.median(kpi_18))}'+W)

                        outSheet.write(z + 1, 18 , float(statistics.median(kpi_18)))

                        print(Y+'uplink_cell_throghput(kbit/s+)'+W, f'= {kpi_19}'+G,
                            f'Median : {float(statistics.median(kpi_19))}'+W)

                        outSheet.write(z + 1, 19 , float(statistics.median(kpi_19)))

                        print(Y+'max_no_user'+W, f'= {kpi_20}'+G,
                            f'Median : {float(statistics.median(kpi_20))}'+W)

                        outSheet.write(z + 1, 20 , float(statistics.median(kpi_20)))

                        print(Y+'number_of_available_downlink_prbs_cell'+W, f'= {kpi_21}'+G,
                            f'Median : {float(statistics.median(kpi_21))}'+W)

                        outSheet.write(z + 1, 21 , float(statistics.median(kpi_21)))

                        print(Y+'average_cqi(huawei_lte_cell)'+W, f'= {kpi_22}'+G,
                            f'Median : {float(statistics.median(kpi_22))}'+W)

                        outSheet.write(z + 1, 22, float(statistics.median(kpi_22)))

                        print(Y+'intra_rat_handover_sr_intra+inter_frequency(huawei_lte_cell'+W, f'= {kpi_23}'+G,
                            f'Median : {float(statistics.median(kpi_23))}'+W)

                        outSheet.write(z + 1, 23, float(statistics.median(kpi_23)))

                        print(Y+'rrc_connection_setup_success_rate_service'+W, f'= {kpi_24}'+G,
                            f'Median : {float(statistics.median(kpi_24))}'+W)

                        outSheet.write(z + 1, 24, float(statistics.median(kpi_24)))

                        print(Y+'s1signal_e-rab_setup_sr(hu_cell)'+W, f'= {kpi_25}'+G,
                            f'Median : {float(statistics.median(kpi_25))}'+W)

                        outSheet.write(z + 1, 25, float(statistics.median(kpi_25)))

                        print(Y+'cell_unvail_duration_daily(huawei_cell_lte)'+W, f'= {kpi_26}'+G,
                            f'Median : {float(statistics.median(kpi_26))}'+W)

                        outSheet.write(z + 1, 26, float(statistics.median(kpi_26)))

                        print(Y+'rssi_pucch(huawei_lte_cell)'+W, f'= {kpi_27}'+G,
                            f'Median : {float(statistics.median(kpi_27))}'+W)

                        outSheet.write(z + 1, 27, float(statistics.median(kpi_27)))

                        print(Y+'rssi_pusch(huawei_lte_cell)'+W, f'= {kpi_28}'+G,
                            f'Median : {float(statistics.median(kpi_28))}'+W)

                        outSheet.write(z + 1, 28, float(statistics.median(kpi_28)))

                        print(Y+'cell_availability_rate_exclude_blocking(cell_hu)'+W, f'= {kpi_29}'+G,
                            f'Median : {float(statistics.median(kpi_29))}'+W)

                        outSheet.write(z + 1, 29, float(statistics.median(kpi_29)))

                        print(Y+'cell_availability_rate_include_blocking(cell_hu)'+W, f'= {kpi_30}'+G,
                            f'Median : {float(statistics.median(kpi_30))}'+W)

                        outSheet.write(z + 1, 30, float(statistics.median(kpi_30)))

                        print(Y+'cell_availability_rate_include_blocking'+W, f'= {kpi_31}'+G,
                            f'Median : {float(statistics.median(kpi_31))}'+W)

                        outSheet.write(z + 1, 31, float(statistics.median(kpi_31)))

                        print(Y+'cell_availability_rate_include_blocking(cell_hu_no_null)'+W, f'= {kpi_32}'+G,
                            f'Median : {float(statistics.median(kpi_32))}'+W)

                        outSheet.write(z + 1, 32, float(statistics.median(kpi_32)))
                        
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

                pass
