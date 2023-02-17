# -*- coding: utf-8 -*-
"""
Created on Thu Feb 16 12:52:34 2023

@author: Admin
"""
import re
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 讀取 Excel 文件
b=input('請輸入日期(YYYY-MM-DD):')
a=input('輸入跌價排名:')

df = pd.read_excel('手機價格.xlsx', sheet_name=b) #選擇目標檔案、工作業
df['漲跌'] = df['漲跌'].replace('%', '', regex=True).astype(float) / 100 #文字轉換成數字

df = df.sort_values(by='漲跌')#選擇列
top_3 = df.head(len(df))[int(a)-1:int(a)+2]#排大小

drop = -top_3['漲跌']       # 跌幅
N = len(-top_3['漲跌'])                  # 計算長度
x = top_3['機型']                # 長條圖x軸座標

width = 0.25                  # 長條圖寬度
plt.bar(x, drop, width)        # 繪製長條圖
plt.xlabel('Product')
plt.title('Range of a price drop')

plt.xticks(x, top_3['機型'])
plt.yticks(np.arange(0, 0.6, 0.05))
plt.show()


































