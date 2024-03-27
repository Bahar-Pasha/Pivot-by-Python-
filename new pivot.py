# -*- coding: utf-8 -*-
"""
Created on Wed Sep 20 16:19:48 2023

@author: B.Pashazanosi
"""

import pandas as pd
import numpy as np 
path=r"C:\Users\B.Pashazanosi\Desktop\Media Pardazesh\گزارش کلی مدیا .xlsx"
df = pd.read_excel(path)
df['Date Invoice'] = pd.to_datetime(df['Date Invoice'], format='%Y/%m/%d')



start_date = pd.to_datetime("2024-03-08")
end_date = pd.to_datetime("2024-03-15")

df= df[(df['Date Invoice'] >= start_date) & (df['Date Invoice'] <= end_date)]

Brand= df['Brand'] =='HONOR'
Category=df['Category'] == 'Mobile'
Region = df['Zone'].isin(["Center~Offline","Center~Online", "East", "South", "West / B2B","West"])
df= df.loc[(Brand)&(Category)& (Region)]


#Combine شهین  and درنا همرا نقش جهان 
df['User Name'] = df['User Name'].replace(9134498006, 9134498004)

#Combine توسعه تجارت آراد همراه and  بازارسازان نامی نت  
df['User Name'] = df['User Name'].replace(9122783620, 9129348939)


# Pivot table 
pivot_table = df.pivot_table(values='Quantity', index=['User Name','Full Name','Zone',"State"], 
                             columns=['Model'], aggfunc='sum')

# Calculate the Grand Total across each row and add it as a new column
pivot_table['Achieved'] = pivot_table.sum(axis=1)

# Sort the pivot table by the 'Region' column
pivot_table_sorted = pivot_table.sort_values(by='Zone')

pivot= r"C:\Users\B.Pashazanosi\Desktop\Media Pardazesh\pivot.xlsx"
pivot_table_sorted.to_excel(pivot)
# Now, let's read the merged DataFrame from the output file to investigate column names
merged_df = pd.read_excel(pivot)
