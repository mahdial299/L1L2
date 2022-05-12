
import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
import statistics

R = '\033[31m'
G = '\033[32m'
C = '\033[36m'
W = '\033[0m'


df = pd.read_excel('test_data.xlsx')

astro = df.to_dict('records')   

# # -------------------------------------
if __name__ == '__main__':

    cell_source = ['A', 'B', 'C']

    for z in range(len(cell_source)):

        kpi_1 = []
        kpi_2 = []
        kpi_3 = []

        for i in range(len(astro)):

            if astro[i]['cell'] == cell_source[z]:

                kpi_1.append(astro[i]['kpi_1'])
                kpi_2.append(astro[i]['kpi_2'])
                kpi_3.append(astro[i]['kpi_3'])

        print(cell_source[z], kpi_1, kpi_2, kpi_3)

       



        


