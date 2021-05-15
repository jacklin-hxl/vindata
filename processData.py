
import pandas as pd
import numpy as np

# 筛选指定天数sku销售数据summary
processDate = [20210401, 20210402, 20210403]
df = pd.read_excel("D:\\tmp\\1.xlsx", sheet_name="4月")
df2 = pd.read_excel("D:\\tmp\\666.xlsx", sheet_name="4月品牌团", usecols=[0,1], dtype='object')

df = df[df["日期"].isin(processDate)]
df = pd.pivot_table(df,index="商品ID",values=["支付金额","支付件数", "商品访客数", "成交人数"], aggfunc=[np.sum])
df.columns = ['商品访客数','成交人数','支付件数','支付金额']
df = df.reset_index()
df2 = df2.dropna(axis=0, how="all")
df2 = pd.pivot_table(df2,index=["商品ID"])
df3 = pd.merge(df2,df,on="商品ID",how='inner')
df3.to_csv
# print(df)

# 求单天总销量
processDate = [20210401]
df = df[df["日期"].isin(processDate)]
df = pd.pivot_table(df,index="商品ID",values=["支付金额"], aggfunc=[np.sum])
df2 = pd.pivot_table(df2,index=["商品ID"])
df3 = df.merge(df2,left_index=True,right_index=True)
print("FDFDF")
