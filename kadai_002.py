#coding:cp932
import pandas as pd

df = pd.DataFrame({"���t":["2023-05-17","2023-05-18","2023-05-19",\
  "2023-05-20","2023-05-21"],"�Ј���":["�R�c","����","���","�c��",\
  "����"],"����":[100,200,150,300,250],"����":["���[�J�[","�㗝�X",\
  "���[�J�[","����","�㗝�X"]})

df["���ϔ���"] = df["����"].mean()
print(df["���ϔ���"])

average_sales = df['����'].mean()
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

df["�Ɛ�"] = df["����"].apply(performance)

writer =  pd.ExcelWriter("�Ɛ�.xlsx")
print(writer)

df.to_excel(writer,sheet_name="Sheet",index = False)

writer.close()