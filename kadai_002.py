#coding:cp932
import pandas as pd

df = pd.DataFrame({"日付":["2023-05-17","2023-05-18","2023-05-19",\
  "2023-05-20","2023-05-21"],"社員名":["山田","佐藤","鈴木","田中",\
  "高橋"],"売上":[100,200,150,300,250],"部門":["メーカー","代理店",\
  "メーカー","商社","代理店"]})

df["平均売上"] = df["売上"].mean()
print(df["平均売上"])

average_sales = df['売上'].mean()
above_averege = average_sales + 50

def performance(level):
  achievement = "";
  if level >= above_averege:
    achievement = "A";
  elif level >= average_sales:
    achievement = "B";
  else:
    achievement ="C" ;
  return achievement

df["業績"] = df["売上"].apply(performance)

writer =  pd.ExcelWriter("業績.xlsx")
print(writer)

df.to_excel(writer,sheet_name="Sheet",index = False)

writer.close()