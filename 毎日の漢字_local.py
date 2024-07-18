import pandas as pd
import numpy as np
from openpyxl import load_workbook

df=pd.read_excel(r'C:\Users\ruben\Repositorios\kanjis\data/漢字.xlsx')
readings = ['次','くんよみ','おんよみ']
aux = ['次','いみ','漢字',]
new_kanji = df.loc[df['新しい']==1]
df = df.loc[df['新しい']==0]
df['random'] =  np.random.random(size=len(df))
df.sort_values(by="random", inplace = True, ascending=False)
daily_kanji_readings = pd.concat([df[readings].head(100).reset_index(drop = True),new_kanji[readings]])
daily_kanji = pd.concat([df[aux].head(100).reset_index(drop = True),new_kanji[aux]])
daily_kanji_readings.to_excel('C:/Users/ruben/Desktop/毎日の漢字.xlsx',"lecturas",index = False)
with pd.ExcelWriter('C:/Users/ruben/Desktop/毎日の漢字.xlsx', engine='openpyxl')  as writer:
    daily_kanji_readings.to_excel(writer, sheet_name = "lecturas",index = False)            
    daily_kanji.to_excel(writer, sheet_name = "漢字, いみ", index=False)
    
    

