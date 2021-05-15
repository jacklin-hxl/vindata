import re
import datetime

import pandas as pd
import numpy as np
from common import getDateList
from collections import Counter


# 筛选指定天数sku销售数据summary
def run(summaryDataFile=None, summaryDataSheet=None, totalDataFile=None, totalDataSheet=None):
    summaryDataFile = summaryDataFile.replace("file:///","")
    totalDataFile = totalDataFile.replace("file:///","")

    summaryDataFile = "D:\\tmp\\666.xlsx"
    summaryDataSheet = "4月品牌团"
    totalDataFile = "D:\\tmp\\1.xlsx"
    totalDataSheet = "2021汇总"

    allTotalData = pd.read_excel(totalDataFile, sheet_name=totalDataSheet, dtype='object')

    allSummaryData = pd.read_excel(summaryDataFile, sheet_name=summaryDataSheet, dtype='object')
    search = ["ID"]
    to_search = re.compile('|'.join(sorted(search, key=len, reverse=True)))
    matches = (to_search.search(el) for el in list(allSummaryData.columns))
    times = Counter(match.group() for match in matches if match)
    times = times["ID"]

    currentCols = [0, 1]
    resultList = []

    for i in range(times):

        summaryData = pd.read_excel(summaryDataFile, sheet_name=summaryDataSheet, usecols=currentCols, dtype='object')
        summaryData.rename(columns={summaryData.columns[0]: "id"},inplace=True)

        oldProcessDate = summaryData.columns[1]
        d1 = re.findall(r"(\d*\.\d*)日",oldProcessDate)
        startDate = str(datetime.datetime.now().year) + "." + d1[0]
        endDate = str(datetime.datetime.now().year) + "." + d1[1]
        dateList = getDateList(start=startDate, end=endDate)

        d2 = re.findall(r"(\d*)点",oldProcessDate)

        isCompare = False
        if '10' and "21" in d2:
            startDate = [dateList[0]]
            endDate =  [dateList[-1]]
            summaryData = summaryData.dropna(axis=0, how="all")

            tmp1 = allTotalData[allTotalData["日期"].isin(startDate)]
            tmp1 = pd.pivot_table(tmp1,index="id", values=["支付金额"], aggfunc=[np.sum])

            tmp2 = allTotalData[allTotalData["日期"].isin(endDate)]
            tmp2 = pd.pivot_table(tmp2, index="id", values=["支付金额"], aggfunc=[np.sum])

            result1 = pd.merge(summaryData, tmp1, on="id",how='left')
            result1.drop(result1.columns[1], axis=1, inplace=True)
            sum1 = result1.sum()[1]

            result2 = pd.merge(summaryData,tmp2, on="id",how='left')
            result2.drop(result2.columns[1], axis=1, inplace=True)
            sum2 = result2.sum()[1]

            if sum1 > sum2:
                dateList.pop(-1)
            else:
                dateList.pop(0)
            isCompare = True
            newProcessDate = oldProcessDate + "(" + str(dateList[0])[4:] + "-" + str(dateList[-1])[4:]  +")"

        totalData = allTotalData[allTotalData["日期"].isin(dateList)]
        totalData = pd.pivot_table(totalData,index="id",values=["支付金额","支付件数", "商品访客数", "成交人数"], aggfunc=[np.sum])
        totalData.columns = ['商品访客数','成交人数','支付件数','支付金额']
        totalData = totalData.reset_index()

        summaryData = summaryData.dropna(axis=0, how="all")
        result = pd.merge(summaryData,totalData,on="id",how='left')
        result = result.reindex(columns=[list(result)[0],list(result)[1],"支付金额", "支付件数","成交人数", "商品访客数", "转化率", "客单价"])

        result['客单价'] = result["支付金额"] / result["成交人数"]
        result['转化率'] = result["成交人数"] / result["商品访客数"]
        summaryValue = ["汇总", "", result.sum()[2], result.sum()[3], result.sum()[4], result.sum()[5], result.sum()[4] / result.sum()[5], result.sum()[2] / result.sum()[4]]
        result.loc[result.shape[0]] = dict(zip(result.columns, summaryValue))

        result['转化率'] = result["转化率"].apply(lambda x: format(x,'.2%'))
        result['客单价'] = result['客单价'].round(decimals=2)
        result['id'] = result['id'].map(str)

        if isCompare:
            result.rename(columns={oldProcessDate: newProcessDate},inplace=True)

        currentCols = [ i + 3 for i in currentCols]
        result = result.fillna(0)
        resultList.append(result)

    writer = pd.ExcelWriter('./tmp.xlsx', mode='a')
    startCol = 0
    for i in resultList:
        i.to_excel(writer, sheet_name=summaryDataSheet, startcol=startCol, index=False)
        startCol += 10
    writer.save()

if __name__ == "__main__":
    run()