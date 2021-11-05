import math

import pandas as pd
import xlsxwriter
import numpy as np
from scipy.stats.contingency import chi2_contingency
from scipy.stats.contingency import margins
import pyreadstat

def write_long_crosttab_to_xslx(df_name, varx, vary, PONDER = 'weight', file_name = 'ispis_baze.xlsx'):
    ''' Takes in spss file and for each specified variable in varx makes crosstab with posthoc based on adjusted standradardized residual with variables in vary
    :param df: spss input
    :param PONDER: weight
    :param file_name: out put file (.xlsx file)
    :return: None
    '''

    df = pd.read_spss(df_name, convert_categoricals=False)
    throwaway, meta = pyreadstat.read_sav(df_name, apply_value_formats=True)

    workbook = xlsxwriter.Workbook(file_name) # opens EXCEL file which is named by the file_name argument

    # defining the cell formating
    cell_format = workbook.add_format() # format regulating method
    cell_format.set_font_color('#000000') # color code
    cell_format.set_font_size(12) # font size
    cell_format.set_align('center') # alignment

    # format for p < 0.01 and standardized residual value higher than 2.58
    cell_formatv1 = workbook.add_format()
    cell_formatv1.set_font_color('#000000') # color code
    cell_formatv1.set_font_size(12) # font size
    cell_formatv1.set_align('center') # alignment
    cell_formatv1.set_bg_color('#00FF80') # bg color code

    # format for p < 0.05 and standardized residual value higher than 1.96
    cell_formatv2 = workbook.add_format()
    cell_formatv2.set_font_color('#000000') # color code
    cell_formatv2.set_font_size(12) # font size
    cell_formatv2.set_align('center') # alignment
    cell_formatv2.set_bg_color('#00FF80') # bg color code

    # format for p < 0.05 and standardized residual value lower than -1.96
    cell_formatm1 = workbook.add_format()
    cell_formatm1.set_font_color('#000000') # color code
    cell_formatm1.set_font_size(12) # font size
    cell_formatm1.set_align('center') # alignment
    cell_formatm1.set_bg_color('#FFCC99') # bg color code

    # format for p < 0.01 and standardized residual value lower than -2.58
    cell_formatm2 = workbook.add_format()
    cell_formatm2.set_font_color('#000000') # color code
    cell_formatm2.set_font_size(12) # font size
    cell_formatm2.set_align('center') # alignment
    cell_formatm2.set_bg_color('#FFCC99') # bg color code

    for column in varx:
        col = 0 # we start from the 1st column
        row = 1 # we start from the second row (because we want to use the 1st row for the "column" variable names)
        worksheet = workbook.add_worksheet(column)

        worksheet.write(row, col, column, cell_format) # we write the "row" variable name
        row += 1
        for item in sorted(df[column].unique()):
            worksheet.write(row, col, meta.variable_value_labels[column][item], cell_format)
            row += 1

        col = 1 # we go to the next column

        for column2 in vary:
            row = 0 # we write the "column" variable name in row 0
            worksheet.write(row, col, column2, cell_format)
            row += 1
            for item in sorted(df[column2].unique()):
                worksheet.write(row, col, meta.variable_value_labels[column2][item], cell_format)
                col += 1

            col -= len(df[column2].unique())
            row += 1

            if column == column2:
                col += len(df[column].unique())
            else:
                crosstab = pd.crosstab(df[column], df[column2], df[PONDER], aggfunc=sum)
                crosstab = crosstab.fillna(0)
                crosstab = crosstab.round(0)
                print(crosstab)
                print(crosstab.columns)
                print(crosstab.index)

                # posthoc format
                residuals_format = chi_square_post_hoc(crosstab)

                for i in range(len(residuals_format)):
                    if residuals_format[i] == 1:
                        residuals_format[i] = cell_formatv1
                    elif residuals_format[i] == 2:
                        residuals_format[i] = cell_formatv2
                    elif residuals_format[i] == 3:
                        residuals_format[i] = cell_formatm2
                    elif residuals_format[i] == 4:
                        residuals_format[i] = cell_formatm1
                    else:
                        residuals_format[i] = cell_format

                residuals_format = np.asarray(residuals_format)
                residuals_format = np.split(residuals_format, crosstab.shape[0])

                crosstab = (100. * crosstab / crosstab.sum()).round(1)
                crosstab = crosstab.to_numpy()
                for i in range(crosstab.shape[0]):
                    for m in range(crosstab.shape[1]):
                        worksheet.write(row, col, f'{crosstab[i][m]}%', residuals_format[i][m])
                        col += 1
                    col -= crosstab.shape[1]
                    row += 1
                col += crosstab.shape[1]


    workbook.close()


def chi_square_post_hoc(a):
    '''
    :param a: Crosstabs to calculate chi square post hoc on
    :return: Table os same shape of Crosstabs table with formatting rules for appropriate cell
    '''
    res = []
    chi, p, dof, expected = chi2_contingency(a, correction= False)
    b = np.asarray(a)
    suma = b.sum()
    b = b.tolist()
    marginss = margins(a)

    if p <= 0.05:
        for i1, item1 in enumerate(b):
            for i2, item2 in enumerate(item1):
                azr = (b[i1][i2] - expected[i1][i2]) / math.sqrt(marginss[0][i1][0] * marginss[1][0][i2] * (1 - marginss[0][i1][0]/suma) * (1 - marginss[1][0][i2]/suma) / suma)
                if b[i1][i2] == 0:
                    res.append(5)
                elif azr >= 2.58:
                    res.append(1)
                elif azr >= 1.96:
                    res.append(2)
                elif azr <= -2.58:
                    res.append(3)
                elif azr <= -1.96:
                    res.append(4)
                else:
                    res.append(5)
    else:
        for dim in b:
            for item in dim:
                res.append(5)

    return res

if __name__ == '__main__':
    # write_long_crosttab_to_xslx('imematrice', [var_row_1, var_row_2, var_row_3], ['var_colum_1', 'var_column_2','var_colum_3'])
    pass