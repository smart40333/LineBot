import pandas as pd
import pyodbc

def strip_whitespace(df): #去除所有欄位的前後空白    
    return df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

conn = pyodbc.connect('DSN=PSCDB', UID='FSP631', PWD='FSP631')
connSQL = pyodbc.connect(driver = 'ODBC Driver 18 for SQL Server', server = '10.72.228.139', user ='sa', password = 'Self@pscnet', database = 'CBAS', TrustServerCertificate='yes')

df_cus = strip_whitespace(pd.read_sql("select CUSID, CUSNAME from FSPFLIB.FSPCS0M WHERE CBASCODE = 'Y'", conn))
df_Line = pd.read_sql("select * from LineID", connSQL)

df_con = pd.merge(df_cus, df_Line, on='CUSID', how='left')
df_con.rename(columns={'CUSNAME_x': 'CUSNAME'}, inplace=True)
df_con = df_con[['CUSID', 'CUSNAME', 'LineID']]

df_con.to_csv('C:\Py_Project\LineBot\mapping.csv', index=False, encoding='utf-8-sig', header=True)

print(df_con)








