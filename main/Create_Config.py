import pandas as pd
import os
config={'filepath':'spreadsheets/spreadsheet.ods',
        'appname':'WhatsApp',
        'sheetpath':'c',
        'pdftoread':'configuration/the-1611-king-james-apocrypha.pdf',
        'click1':[0,0],
        'click2':[0,0],
        'click3':[0,0],
        'click4':[0,0],
        'testcall':''}
config.update(click1=[1,1])
df=pd.DataFrame(config)
df.to_json('configuration\config.json',orient='index')
df2=pd.read_json('configuration\config.json',orient='index')
print(df2)
#df.to_json('data.json',orient='table')