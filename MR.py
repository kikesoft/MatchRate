#GLOBAL----------------------------------------------------------------------------------
import pandas as pd
import xlsxwriter
import numpy as np
df = pd.read_csv("C:\\Users\\enrique.silva\\Documents\\Python Scripts\\RealData.csv")


#DEFINITIONS-----------------------------------------------------------------------------
TotalColumns = len(df.columns)
TotalRows = len(df)
AllNames = []
Threshold = 85
index=0
for i in df.columns:
    AllNames.append(df.columns[index])
    index+=1

w, h = TotalRows, TotalRows

Matrix = [[0 for x in range(w)] for y in range(h)]
FullData = [[0 for x in range(w)] for y in range(h)]
df = df.replace(np.nan, "NORESPONSE")


#MATCH RATE CALCULATION ------------------------------------------------------------------
x = 0
for k in df["respid"]:
    if x<TotalRows:
        y = x
        for j in df["respid"]:
            if y<TotalRows:
                cont = 0
                tr = 0
                index = 0
                for i in AllNames[1:]:
                    if not pd.isna(df[i][x]):
                        tr+=1
                        if df[i][x] == df[i][y]:
                            cont+=1
                Match = round((cont / tr)*100,2)
                Matrix[x][y] = Match
            y+=1
        x+=1

#OUTPUT SECTION------------------------------------------------------------------------------
output = "C:\\Users\\enrique.silva\\Documents\\Python Scripts\\MR\\8509MR.xlsx"
workbook = xlsxwriter.Workbook(output)
worksheet = workbook.add_worksheet("MatchRates")
worksheet2 = workbook.add_worksheet("Highest")
x = 1
y = 0
for i in df["respid"]:
    worksheet.write(x, y, i)
    worksheet.write(y, x, i)
    x+=1

index = 0
for i in Matrix:
    x=1
    y=index+1
    for j in Matrix[index]:
        worksheet.write(y, x, j)
        x+=1
    index+=1


for i in range(len(Matrix[0])):
    if i < TotalRows:
        index = i + 1
        cont = i + 1
        for j in Matrix[i][i:]:
            #worksheet.write(col, row, data)
            worksheet.write(index, cont , j)
            index+=1
        cont+=1



# Add a format. Light red fill with dark red text.
format1 = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

# Add a format. Green fill with dark green text.
format2 = workbook.add_format({'bg_color': '#CCFF00',
                               'font_color': '#006100'})



worksheet.conditional_format('B2:AAA9999', {'type': 'cell',
                                         'criteria': '>=',
                                         'value': 70,
                                         'format': format1})


worksheet.conditional_format('B2:AAA9999', {'type': 'cell',
                                         'criteria': '<',
                                         'value': 70,
                                         'format': format2})


#FILLING INFO BY RESPID---------------------------------------------------------------------------------


df2 = pd.read_excel (output)


for j in range(TotalRows):
    k = j + 1
    for i in range(TotalRows):
        index = i + 1
        FullData[j][i] = df2.iloc[:,k][i]

auxlen = len(FullData[0])
MaxRates = [0 for x in range(auxlen)]
MaxRespid = [0 for x in range(auxlen)]
for i in range(auxlen):
    if i < auxlen:
        aux = FullData[i]
        aux = np.array(aux)
        MaxRates[i] = np.amax(aux[aux<100])
index = 0
for i in MaxRates:
    MaxRespid[index] = df["respid"][FullData[index].index(i)]
    index+=1

worksheet2.write(0, 0, "Respid")
worksheet2.write(0, 1, "MaxRate")
worksheet2.write(0, 2, "With Respid:")
worksheet2.write(0, 3, "Type:")

index = 0
for i in df["respid"]:
    x=index+1;
    d1 = df["respid"][index]
    d2 = MaxRates[index]
    d3 = MaxRespid[index]
    if d2 > Threshold:
        d4 = "FRAUD"
    else:
        d4 = "OK"
    worksheet2.write(x, 0, d1)
    worksheet2.write(x, 1, d2)
    worksheet2.write(x, 2, d3)
    worksheet2.write(x, 3, d4)
    index+=1


workbook.close()
