from re import A
import pandas as pd

lichbay = pd.read_excel('lichbay.xlsx', sheet_name="6")
A320 = pd.read_excel('doi1.xlsx', sheet_name="A320")
A350 = pd.read_excel('doi1.xlsx', sheet_name="A350")
ATR = pd.read_excel('doi1.xlsx', sheet_name="ATR")
count_A320 = 0
count_A350 = 0
count_ATR = 0
df = lichbay
df['Nhan vien phu trach'] = ""
print(A320.iat[count_A320,0])
for i in range (len(df.index)):
    if df['AcType'][i] == "A320" or df['AcType'][i] == "A321":
        if count_A320 < len(A320):
            df['Nhan vien phu trach'][i] = A320.iat[count_A320,0]
            count_A320 = count_A320 + 1
        else:
            count_A320 = 0
            df['Nhan vien phu trach'][i] = A320.iat[count_A320,0]
            count_A320 = count_A320 + 1
    elif df['AcType'][i] == "A350":
        if count_A350 < len(A350):
            df['Nhan vien phu trach'][i] = A350.iat[count_A350,0]
            count_A350 = count_A350 + 1
        else:
            count_A350 = 0
            df['Nhan vien phu trach'][i] = A350.iat[count_A350,0]
            count_A350 = count_A350 + 1
    else:
        if count_ATR < len(ATR):
            df['Nhan vien phu trach'][i] = ATR.iat[count_ATR,0]
            count_ATR = count_ATR + 1
        else:
            count_ATR = 0
            df['Nhan vien phu trach'][i] = ATR.iat[count_ATR,0]
            count_ATR = count_ATR + 1
print(df['Nhan vien phu trach'])
df.to_excel("output.xlsx")



# df2 = pd.read_excel('DOI 1.xlsx')
# print(df2[])