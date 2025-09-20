import pandas as pd
import pyexcel as pe
import random
import string

length = 2
a=[]
b=[]
c=[]
d=[]
e=[]
f=[]
testnumber=f"xxxxxxx"
for i in range(5):
    length = 4
    random_string = ''.join(random.choices(string.ascii_letters + string.digits, k=length))
    print(random_string)
    d.append(random_string)
    a.append(testnumber)
    b.append('')
    c.append(10)
    e.append('')
    if ((i%2)==0):
        x=True
    else: 
        x=False
    f.append(x)
data = {
    'Account-Name': d,
    'Phone-Number': a,
    'Status': b,
    'Call-File':e,
}
data2={
    'Account-Name': d,
    'Current-Page':c,
    'Reading':f,
}
df = pd.DataFrame(data)
print(df)
df.to_csv('spreadsheets\spreadsheet.csv')
df2= pd.DataFrame(data2)
df2.to_csv('configuration\\acc-log.csv')
# Read CSV data
records = pe.get_records(file_name="spreadsheets\spreadsheet.csv")
# Save records to an ODS file
pe.save_as(records=records, dest_file_name="spreadsheets\spreadsheet.ods")
records2 = pe.get_records(file_name="configuration\\acc-log.csv")
# Save records to an ODS file
pe.save_as(records=records2, dest_file_name=f"configuration\\acc-log.ods")