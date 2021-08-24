from os import wait
import numpy as np
import pandas as pd
from pandas.core.frame import DataFrame
from pandas.core.series import Series;

mainDf = pd.read_excel('classifyVA_NVA.xlsx')
nvaDf = pd.read_excel('NVA.xlsx')
svaDf = pd.read_excel('SVA.xlsx')
vaDf = pd.read_excel('VA.xlsx')

SVAfilteredDf = mainDf[mainDf.Category == 'SVA']
VAfilteredDf = mainDf[mainDf.Category == 'VA']
NVAAfilteredDf = mainDf[mainDf.Category == 'NVA']

finalSVADf = DataFrame()
for col in svaDf.columns:
    for val in svaDf[col].to_list():
        resDf = DataFrame()
        if str(val) != 'nan':
            resDf = SVAfilteredDf[SVAfilteredDf.Description.str.contains(val, case=False)]
            resDf = resDf.assign(Classification=col)
        if not resDf.empty:
            finalSVADf = pd.concat([finalSVADf,resDf])
            
# print('SVA Dataframe: ')
# print(finalSVADf)

finalVADf = DataFrame()
for col in vaDf.columns:
    for val in vaDf[col].to_list():
        resDf = DataFrame()
        if str(val) != 'nan':
            resDf = VAfilteredDf[VAfilteredDf.Description.str.contains(val, case=False)]
            resDf = resDf.assign(Classification=col)        
        if not resDf.empty:
            finalVADf = pd.concat([finalVADf,resDf])

# print('\n\nVA Dataframe: ')
# print(finalVADf)

finalNVADf = DataFrame()
for col in nvaDf.columns:
    for val in nvaDf[col].to_list():
        resDf = DataFrame()
        if str(val) != 'nan':
            resDf = NVAAfilteredDf[NVAAfilteredDf.Description.str.contains(val, case=False)]
            resDf = resDf.assign(Classification=col)        
        if not resDf.empty:
            finalNVADf = pd.concat([finalNVADf,resDf])

# print('\n\nNVAA Dataframe: ')
# print(finalNVADf)

finalDf = DataFrame()
finalDf = pd.concat([finalDf,finalSVADf])
finalDf = pd.concat([finalDf,finalVADf])
finalDf = pd.concat([finalDf,finalNVADf])

finalDf = pd.concat([finalDf,mainDf])

finalDf = finalDf[~finalDf.index.duplicated(keep='first')]
finalDf = finalDf.sort_index(axis=0)

writer = pd.ExcelWriter("result.xlsx", engine='xlsxwriter')
finalDf.to_excel(writer, index=False)
writer.save()