import pyodbc
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from pandas import read_excel

'''
This updates the ESD(Expected Ship By Date) that Dell sends on the WO to the calculated - 
MSBD(Must Ship By Date) that is sent via their disti/retail file. 
'''

Tk().withdraw()
    
esdFile = askopenfilename()
user = input('Enter PkMS User ID: ')
pwd = input('Enter PkMS Password: ')
driver = "iSeries Access ODBC Driver"
conn = pyodbc.connect(f"Driver={driver};System=PKMS.US.LOGISTICS.CORP;Uid={user};Pwd={pwd};")
cursor = conn.cursor()
df = read_excel(esdFile, engine='openpyxl')
i = 0


for row in df.itertuples():
    ctrl_num = str(row[1])
    esd = str(row[2])
    i += 1

    cursor.execute(f"UPDATE CAPM01.WM0272PRDD.PHPICK00 SET PHCMDT = '{esd}' WHERE PHPCTL = '{ctrl_num}'")
    conn.commit()
 

print(f"{i} records have been updated.")
