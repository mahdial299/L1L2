
import matplotlib.pyplot as plt
import os




df_hu = pd.read_excel(f'HU_Efficiency_Data{yestr_y}{yestr_m}{yestr_d}.xlsx', sheet_name = f'{item}')


x_data = df_hu[['User_MHZ']]
y_data = df_hu[['DL_User_Throughput']]
pro_data = df_hu[['pro']]
sec_data = df_hu[['SECTOR']]


x_data = np.asanyarray(x_data)
y_data = np.asanyarray(y_data)
pro_data = np.asanyarray(pro_data)
sec_data = np.asanyarray(sec_data)


x_data = x_data.flatten()
y_data = y_data.flatten()
pro_data = pro_data.flatten()
sec_data = sec_data.flatten()



x_data = np.nan_to_num(x_data)
y_data = np.nan_to_num(y_data)

x_data = list(x_data)
y_data = list(y_data)
pro_data = list(pro_data)
sec_data = list(sec_data)


# =========================================== check from here 

province_list = [
'AR',
'QN',
'ZN',
'BU',
'HZ',
'LN',
'KD',
'KS',
'HN',
'FS',
'ES',
'AG',
'AS',
'IL',
'KZ',
'TH',
'KM',
'QM']

for z in range(len(province_list)):

    x2 = []
    y2 = []
    AG_sector = []
    

    for i in range(len(pro_data)):

        if pro_data[i] == province_list[z]:

            x2.append(x_data[i])
            y2.append(y_data[i])
            AG_sector.append(sec_data[i])
            

    
    plt.scatter(x2, y2, color = 'yellow')

    # -------------------------------- worst cells

    x3 = []
    y3 = []
    AG_worst = []
    pro_worst = []

    
    pro_good = []


    list_status = []

    bh = []


    for i in range(len(pro_data)):
        if pro_data[i] == province_list[z]:

            if y_data[i] < n_a / (n_d * x_data[i] + n_c * (x_data[i] ** 2) + n_b * (x_data[i] ** 3) + n_e):

                AG_worst.append(sec_data[i])
                x3.append(x_data[i])
                y3.append(y_data[i])
                pro_worst.append(pro_data[i])
                list_status.append("WORST")
                bh.append(item)

        
            elif y_data[i] > n_a / (n_d * x_data[i] + n_c * (x_data[i] ** 2) + n_b * (x_data[i] ** 3) + n_e):

                AG_worst.append(sec_data[i])
                x3.append(x_data[i])
                y3.append(y_data[i])
                pro_good.append(pro_data[i])
                list_status.append("GOOD")
                bh.append(item)

                
    print(R + f'province :' + W + f'{dict(Counter(pro_worst))}')
    print(G + f'province :' + W + f'{dict(Counter(pro_good))}')
    # print(len(AG_worst))
    plt.scatter(x3, y3, color = 'red')
    # plt.scatter(x4, y4, color = 'blue')

    plt.plot(x_line, y_line, '--', color='green',
        label=f'Baseline_{item}', linewidth=4)
    
    

    diff_worst = []

    # diff_good = []
    

    for i in range(len(AG_worst)):


        diff_worst.append(y3[i] - (n_a / (n_d * x3[i] + n_c * (x3[i] ** 2) + n_b * (x3[i] ** 3) + n_e)))

    
    os.chdir(fr'{file_directory}\{item}')

    outWorkbook1 = xlsxwriter.Workbook(str(province_list[z])+f"_{item}.xlsx")
    outSheet1 = outWorkbook1.add_worksheet()

    outSheet1.write("A1", "CELLS")
    outSheet1.write(0, 0, "SECTORS")
    # outSheet1.write(0, 0, "Worst SECTORS")
    outSheet1.write(0, 1, "BH time")
    outSheet1.write(0, 2, "User per MHz")
    outSheet1.write(0, 3, "User throughput")
    outSheet1.write(0, 4, "DISTANCE TO EXPECTED THROUGHPUT")
    outSheet1.write(0, 5, "STATUS")


    for k in range(len(AG_worst)):
        outSheet1.write(k + 1, 0, AG_worst[k])
    for q in range(len(bh)):
        outSheet1.write(q + 1, 1, bh[q])
    for j in range(len(x3)):
        outSheet1.write(j + 1, 2, x3[j])
    for k in range(len(y3)):
        outSheet1.write(k + 1, 3, y3[k])
    for m in range(len(diff_worst)):
        outSheet1.write(m + 1, 4, diff_worst[m])
    for o in range(len(list_status)):
        outSheet1.write(o + 1, 5, list_status[o])
    outWorkbook1.close()
# ----------------------------------------------------------------------------------------------------------------

#     plt.style.use('bmh')

# plt.show()

os.chdir(fr'{file_directory}\{item}')
outWorkbook2 = xlsxwriter.Workbook(f"2_Baseline_CELLS_{item}.xlsx")
outSheet2 = outWorkbook2.add_worksheet()
outSheet2.write(0, 0, "X")
outSheet2.write(0, 1, "Y")
for k in range(len(x_line)):
    outSheet2.write(k,0, x_line[k])
for k in range(len(y_line)):
    outSheet2.write(k,1, y_line[k])
outWorkbook2.close()

    # for ergo in tqdm(range(10), colour='green', desc=f'hour {item} progress : '):
    #     time.sleep(1)
