from pprint import pprint

import pymupdf
import pandas as pd
import openpyxl


res = []
file_name = input("Enter file name: ")
doc = pymupdf.open(f"files/input/{file_name}") #  open document
for page in doc:
    tabs = page.find_tables() #  locate and extract any tables on page
    print(f"{len(tabs.tables)} found on {page}") #  display number of found tables
    if tabs.tables:  #  at least one table found ?
       for tab in tabs:
          res.extend(tab.extract())     
          
res_d = {}
res_dict = {}
check = {}
for i, el in enumerate(res[0]):
    if "\n" in el:
        el = el.replace("\n", " ")
    check[i] = el
    res_dict[el] = []
    
for el in res[1:]:
    if None in el:
        continue
    for j, sub_el in enumerate(el):
        if sub_el.replace("\n", " ") != check[j]:
            res_dict[check[j]].append(sub_el.replace("\n", " "))

df = pd.DataFrame.from_dict(res_dict, orient="index")
df = df.transpose()

df.index = df.iloc[:, 0].to_numpy()
df = df.drop(df.columns[0], axis=1)
df.index.name = chr(8470)

df.to_excel(f"files/output/{file_name.rstrip('.pdf')}.xlsx")
