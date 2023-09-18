#GLOBAL----------------------------------------------------------------------------------
import pandas as pd
import xlsxwriter
import numpy as np
df = pd.read_csv("C:\\Users\\enrique.silva\\Documents\\Python Scripts\\J7692.csv")


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
#df = df.replace(np.nan, " ")
