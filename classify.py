import pandas as pd
from pandas.core.frame import DataFrame

mainDf = pd.read_excel('classifyVA_NVA.xlsx')

svaDf = pd.read_excel('SVA.xlsx')
vaDf = pd.read_excel('VA.xlsx')
nvaDf = pd.read_excel('NVA.xlsx')

SVAfilteredDf = mainDf[mainDf.Category == 'SVA']
VAfilteredDf = mainDf[mainDf.Category == 'VA']
NVAAfilteredDf = mainDf[mainDf.Category == 'NVA']

def sortDataFrames(colDf, filteredDf):
    temDf = DataFrame()
    for col in colDf.columns:
        for val in colDf[col].to_list():
            resDf = DataFrame()
            if str(val) != 'nan':
                resDf = filteredDf[filteredDf.Description.str.contains(val, case=False)]
                resDf = resDf.assign(Classification=col)
            if not resDf.empty:
                temDf = pd.concat([temDf,resDf])
    return temDf

finalSVADf = sortDataFrames(svaDf, SVAfilteredDf)
finalVADf = sortDataFrames(vaDf, VAfilteredDf)
finalNVADf = sortDataFrames(nvaDf, NVAAfilteredDf)

finalDf = DataFrame()
finalDf = pd.concat([finalDf,finalSVADf])
finalDf = pd.concat([finalDf,finalVADf])
finalDf = pd.concat([finalDf,finalNVADf])

finalDf = pd.concat([finalDf,mainDf])

finalDf = finalDf[~finalDf.index.duplicated(keep='first')]
finalDf = finalDf.sort_index(axis=0)

print('Saving data to file')

writer = pd.ExcelWriter("result.xlsx", engine='xlsxwriter')
finalDf.to_excel(writer, index=False)
writer.save()

print('Saved data to file')