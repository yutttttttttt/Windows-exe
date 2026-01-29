#!/usr/bin/env python
# coding: utf-8

# In[22]:


import pandas as pd
from openpyxl import load_workbook
import os
import re


# In[23]:


path = 'C:\\Users\\13476\\Desktop\\'
#path = '/Users/clarazhou/Desktop/'
folder_path = path +'check file'
xlsx_path = path +'check excel.xlsx'


# In[24]:


AMOUNT_PATTERN = r"\d+\.\d{2}"
DATE_PATTERN = r"\d{4}"
CHINESE_PATTERN = r"[\u4e00-\u9fff]+"

def rename_ur(file_path, filename):
    new_name = "UR-" + filename
    os.rename(file_path, os.path.join(folder_path, new_name))

df = pd.read_excel(xlsx_path)
print('Reading')
# Amount
df["Amount"] = (
    df["Amount"]
      .astype(str)
      .str.replace("$", "", regex=False)
      .str.replace(",", "", regex=False)
    .astype(float)
      .map(lambda x: f"{x:.2f}")
)
if 'Number_Flag' not in df.columns:
    
        #  Posting Date 
    df = df.sort_values(by = "Posting Date")
    df["Trans Date"] = pd.to_datetime(df["Trans Date"]).dt.strftime("%m%d")
    
    df['Number_Flag'] = range(1, len(df) + 1)
    df["Valid_Flag"] = False


df["Trans Date"] = df["Trans Date"].astype(str).str.zfill(4)
print('Number Flag adding')



# In[25]:


for filename in os.listdir(folder_path):
    if filename.startswith("OK-"):
        continue
    file_path = os.path.join(folder_path, filename)
    if not os.path.isfile(file_path):
        continue

    result_name = filename
    date_value = None
    amount_value = None
    chinese_title = None

    # ---------- Step 1: must be PDF ----------
    if not filename.lower().endswith(".pdf"):
        rename_ur(file_path, filename)
        
        continue

    temp_name = filename
    # ---------- Step 2: remove symbols ----------
    for s in ["USD", "usd", "美金", "$", "-", "_", " ",'UR']:
        temp_name = temp_name.replace(s, "")
    # ---------- Step 3: must start with 3918 ----------
    if temp_name.startswith("3918卡"):
        temp_name = temp_name.replace("3918卡", "", 1)
    elif temp_name.startswith("3918"):
        temp_name = temp_name.replace("3918", "", 1)

    else:
        rename_ur(file_path, filename)
        
        continue

    # ---------- Step 4: must contain amount ----------
    amount_match = re.search(AMOUNT_PATTERN, temp_name)
    if not amount_match:
        rename_ur(file_path, filename)
        
        continue

    amount_value = amount_match.group()
    temp_name = temp_name.replace(amount_value, "", 1)

    # ---------- Step 5: optional date ----------
    date_match = re.search(DATE_PATTERN, temp_name)
    if date_match:
        date_value = date_match.group()
        temp_name = temp_name.replace(date_value, "", 1)

    

    # ---------- Step 6: extract Chinese ----------
    chinese_list = re.findall(CHINESE_PATTERN, temp_name)
    chinese_title = "".join(chinese_list)

    if not chinese_title:
        rename_ur(file_path, filename)
        

        continue
    
    # ---------- Step 7: add Valid_Flag, number_flag, Chinese comments----------
    if date_value:
        mask = (
            (df["Amount"] == amount_value) &
            (df["Trans Date"] == date_value)
        )
    else:
        mask = (df["Amount"] == amount_value)
    
    # create Valid_Flag
    df.loc[mask, "Valid_Flag"] = True
    
    #add Chinese comment
    df.loc[mask, "MCC Description"] = (
        df.loc[mask, "MCC Description"].astype(str)
        + " "
        + str(chinese_title)
    )
    #add number flag
    if mask.any():
        number_flag = int(df.loc[mask, "Number_Flag"].iloc[0])
        new_name = f"OK-{number_flag}-{filename}"
        os.rename(file_path, os.path.join(folder_path, new_name))
        
        continue
    else:
        number_flag = None
print('Saved')        

df.to_excel(xlsx_path, index=False)

input('Press Enter to exit...')
# In[ ]:




