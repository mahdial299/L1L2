from black import main
from torch import le
import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
import statistics
import xlrd

R = '\033[31m'
G = '\033[32m'
C = '\033[36m'
W = '\033[0m'

ozalid = 'calculating kpi median begin...'


# --------------------- test list data

df = pd.read_excel('test_data.xlsx')

astro = df.to_dict('records')   # list of dictionaries

# # -------------------------------------
if __name__ == '__main__':

#    print(astro)
    cell_source = ['A', 'B', 'C']

    for i in range(len(astro)):


        lisk_1 = []

        for z in range(len(cell_source)):
        
            if astro[i]['cell'] == cell_source[z]:

                lisk_1.append(astro[i]['kpi_1'])

        print(astro[i]['cell'], lisk_1)



        


