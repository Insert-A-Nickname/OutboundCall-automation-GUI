from pypdf import PdfReader
import pyttsx3
import pandas as pd
df=pd.read_csv("configuration\\acc-log.csv",index_col=0)
df2=pd.read_excel("spreadsheets\spreadsheet.ods")
book=pd.read_json("configuration\config.json",orient="index").loc[0,"pdftoread"]
print(book)
reader = PdfReader(book)
p=df["Current-Page"][df["Account-Name"]==df2["Account-Name"][0]][0]
print(p)
page = reader.pages[10]
engine = pyttsx3.init()
text = page.extract_text()
engine.setProperty("rate", 200)
engine.say(text)
try:
    engine.runAndWait()
except KeyboardInterrupt:
    engine.stop()
    p=p+1
    df["Current-Page"][df["Account-Name"]==df2["Account-Name"][0]][0]=p
    df.to_csv("configuration\\acc-log.csv")
    print(p)


