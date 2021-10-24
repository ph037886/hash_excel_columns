# -*- coding: utf-8 -*-
# python 3.8.5
from tkinter import Tk #導入選擇檔案功能
from tkinter import filedialog
import pandas as pd
import hashlib
import random
import os

def read_file_name():#讀取舊檔路徑
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(initialdir=r'\file')
    basename = os.path.basename(file_path)
    file_name = os.path.splitext(basename)[0]
    return file_path,file_name

def read_xlsx(path):#讀取Excel
    xlsx=pd.ExcelFile(path)
    sheet_list=xlsx.sheet_names #產生sheet名稱的list供後面使用
    xlsx_list=[]
    for i in sheet_list: #把每個sheet內容讀出來，建成一個list
        temp=xlsx.parse(i,dtype=str)
        xlsx_list.append(temp)
    return sheet_list,xlsx_list

def hash_salt(a,salt):
    m=hashlib.sha256()
    b=(str(a)+str(salt))
    m.update(b.encode('utf-8'))
    h=m.hexdigest()
    return h

filename=read_file_name()
old_xlsx=filename[0] #全路徑
xlsx_name=filename[1] #檔名
temp=read_xlsx(old_xlsx) #讀取Excel開始
sheetname=temp[0]
xlsx_list=temp[1]
salt=random.random() #用亂數加鹽
new_filename=filedialog.asksaveasfilename(defaultextension="*.xlsx",filetypes=(("Excel檔", "*.xlsx"),("純文字檔案","*.csv"),("All files", "*.*")),initialfile=xlsx_name) #存檔路徑，檔名預設為舊檔名
save_xlsx=pd.ExcelWriter(new_filename) #先形成writer

i=0
while i<len(xlsx_list): #依序開啟sheet的資料
    sheet=xlsx_list[i]
    print(sheet.columns) #印出sheet的名稱供參考
    column_number2=[]
    column_number=[str(input("請輸入進行hash的column編號："))] #輸入要做hash的欄位
    while True:
        if column_number2=='Y': #最後一個輸入Y結束
            del column_number[-1]
            break
        else:
            column_number2=str(input("請輸入進行hash的column編號，若需結束請輸入Y：")).upper() #強制大寫
            column_number.append(column_number2)
    for j in column_number: #開始做hash
        j=int(j)
        k=0
        while k<len(sheet):
            sheet.iloc[k,j]=hash_salt(sheet.iloc[k,j],salt)
            k+=1
    sheet.to_excel(save_xlsx,sheet_name=sheetname[i],index=False,freeze_panes=(1,0))
    i+=1
save_xlsx.save()
