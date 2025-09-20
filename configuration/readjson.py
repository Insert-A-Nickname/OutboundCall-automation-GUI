import pandas as pd
df=pd.read_json('configuration/config.json',orient='index')
print(df)
print(df.values)