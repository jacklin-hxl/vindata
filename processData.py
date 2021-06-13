import re
import datetime

import pandas as pd
import numpy as np
from common import getDateList
from collections import Counter


# 筛选指定天数sku销售数据summary
def run(summaryDataFile=None, summaryDataSheet=None, totalDataFile=None, totalDataSheet=None, currentDate=None):
    summaryDataFile = summaryDataFile.replace("file:///","")
    totalDataFile = totalDataFile.replace("file:///","")

    # summaryDataFile = "D:\\tmp\\666.xlsx"
    # summaryDataSheet = "4月品牌团"
    # totalDataFile = "D:\\tmp\\1.xlsx"
    # totalDataSheet = "2021汇总"
    if currentDate == "":
        currentDate = str(datetime.datetime.now().year)

    allTotalData = pd.read_excel(totalDataFile, sheet_name=totalDataSheet, dtype='object')

    allSummaryData = pd.read_excel(summaryDataFile, sheet_name=summaryDataSheet, dtype='object')
    search = ["id"]
    to_search = re.compile('|'.join(sorted(search, key=len, reverse=True)))
    matches = (to_search.search(el) for el in list(allSummaryData.columns))
    times = Counter(match.group() for match in matches if match)
    times = times["id"]

    currentCols = [0, 1]
    resultList = []

    for i in range(times):
        summaryData = pd.read_excel(summaryDataFile, sheet_name=summaryDataSheet, usecols=currentCols, dtype='object')
        summaryData.rename(columns={summaryData.columns[0]: "id"},inplace=True)

        oldProcessDate = summaryData.columns[1]
        d1 = re.findall(r"(\d*\.\d*)日",oldProcessDate)
        startDate = currentDate + "." + d1[0]
        endDate = currentDate + "." + d1[1]
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
        result = result.reindex(columns=[list(result)[0],list(result)[1],"支付金额", "支付件数", "到手价", "商品访客数", "转化率", "客单价", "成交人数", "UV价值"])
        
        result[result.columns[1]] = result[result.columns[1]].map(str)
        result['客单价'] = result["支付金额"] / result["成交人数"]
        result['转化率'] = result["成交人数"] / result["商品访客数"]
        result["到手价"] = result["支付金额"] / result["支付件数"]
        result["UV价值"] = result["支付金额"] / result["商品访客数"]

        resultSum = result.sum()
        totalAmount = resultSum[2]
        totalNumber = resultSum[3]
        averagePrice = totalAmount / totalNumber
        visitors = resultSum[5]
        conversionRate = resultSum[-2] / resultSum[5]
        unitPrice = resultSum[2] / resultSum[-2]
        payNumber = resultSum[-2]
        UVWorth = totalAmount / visitors

        summaryValue = ["汇总", "", totalAmount, totalNumber, averagePrice, visitors, conversionRate, unitPrice, payNumber, UVWorth]
        result.loc[result.shape[0]] = dict(zip(result.columns, summaryValue))

        result['转化率'] = result["转化率"].apply(lambda x: format(x,'.2%'))
        result['客单价'] = result['客单价'].round(decimals=2)
        result["UV价值"] = result["UV价值"].round(decimals=2)
        result["到手价"] = result["到手价"].round(decimals=2)
        result['id'] = result['id'].map(str)

        if isCompare:
            result.rename(columns={oldProcessDate: newProcessDate},inplace=True)

        currentCols = [ i + 3 for i in currentCols]
        result = result.fillna(0)
        resultList.append(result)

    writer = pd.ExcelWriter('./tmp.xlsx', mode="a")
    startCol = 0
    for i in resultList:
        i.to_excel(writer, sheet_name=summaryDataSheet, startcol=startCol, index=False)
        startCol += 12
    writer.save()


def runTow(summaryDataFile=None, summaryDataSheet=None, totalDataFile=None, totalDataSheet=None, yDataFile=None, yDataSheet=None):
    # summaryDataFile = summaryDataFile.replace("file:///","")
    # totalDataFile = totalDataFile.replace("file:///","")
    # yDataFile = yDataFile.replace("file:///","")

    summaryDataFile = "D:\\tmp\\数据源.xlsx"
    summaryDataSheet = "1"
    totalDataFile = "D:\\tmp\\2021年维达原始数据.xlsx"
    totalDataSheet = "2021汇总"
    yDataFile = "D:\\tmp\\2020年维达原始数据.xlsx"
    yDataSheet = "2020汇总"

    sumDataSheet = summaryDataSheet + "汇总"
    currentDate = str(datetime.datetime.now().year)

    LDate = [1, 3, 5, 7, 8, 10, 12]
    currentCols = [0, 1, 2]
    summaryData = pd.read_excel(summaryDataFile, sheet_name=summaryDataSheet, usecols=currentCols, dtype='object')
    summaryData.rename(columns={summaryData.columns[1]: "id"},inplace=True)
    summaryData.fillna(method='pad')

    times = []

    d1 = re.findall(r"(\d*\.\d*)", summaryData.columns[2])
    startDate = currentDate + "." + d1[0]
    endDate = currentDate + "." + d1[1]

    cDataList = getDateList(start=startDate, end=endDate)
    times.append((cDataList, "当月"))

    if d1[0].split(".")[0] == "1":
        mStartDate = currentDate + "." + "1." + d1[0].split(".")[1]
    else:
        mStartDate = currentDate + "." + str(int(d1[0].split(".")[0]) - 1) + "." + d1[0].split(".")[1]    

    if d1[1].split(".")[0] == "1":
        mEndDate = currentDate + "." + "1." + d1[1].split(".")[1]
    else:
        mEndDate = currentDate + "." + str(int(d1[1].split(".")[0]) - 1) + "." + d1[1].split(".")[1]
    if d1[1].split(".")[1] == "31":
        if int(d1[1].split(".")[0]) - 1 not in LDate:
            mEndDate = currentDate + "." + str(int(d1[1].split(".")[0]) - 1) + "." + "30"
    mDateList = getDateList(start=mStartDate, end=mEndDate)
    times.append((mDateList, "上月"))

    yStartDate = str(int(currentDate) - 1) + "." + d1[0]
    yEndDate = str(int(currentDate) - 1) + "." + d1[1]
    yDateList = getDateList(start=yStartDate, end=yEndDate)
    times.append((yDateList, "去年"))

    resultList = []
    totalResultList = []

    for i in times:

        if i[1] == "去年":
            totalDataFile = yDataFile
            totalDataSheet = yDataSheet

        allTotalData = pd.read_excel(totalDataFile, sheet_name=totalDataSheet, dtype='object')

        totalData = allTotalData[allTotalData["日期"].isin(i[0])]
        totalData = pd.pivot_table(totalData,index="id",values=["支付金额","支付件数", "商品访客数", "成交人数"], aggfunc=[np.sum])
        totalData.columns = ['商品访客数','成交人数','支付件数','支付金额']
        totalData = totalData.reset_index()


        summaryData = summaryData.dropna(axis=0, how="all")
        result = pd.merge(summaryData,totalData,on="id",how='left')
        result = result.reindex(columns=[list(result)[0],list(result)[1], list(result)[2], "支付金额", "支付件数", "到手价", "商品访客数", "转化率", "客单价", "成交人数", "UV价值"])
        oldName = list(result)[2]
        newName = oldName + "(" + i[1] + ")"
        result.rename(columns={oldName: newName},inplace=True)

        result[result.columns[1]] = result[result.columns[1]].map(str)
        result['客单价'] = result["支付金额"] / result["成交人数"]
        result['转化率'] = result["成交人数"] / result["商品访客数"]
        result["到手价"] = result["支付金额"] / result["支付件数"]
        result["UV价值"] = result["支付金额"] / result["商品访客数"]
        result["品系"] = result["品系"].fillna(method='pad')

        # result['转化率'] = result["转化率"].apply(lambda x: format(x,'.2%'))
        result['转化率'] = result["转化率"].round(decimals=2)
        result['转化率'] = result["转化率"].replace(np.nan, 0)
        result['客单价'] = result['客单价'].round(decimals=2)
        result["UV价值"] = result["UV价值"].round(decimals=2)
        result["到手价"] = result["到手价"].round(decimals=2)
        result['id'] = result['id'].map(str)

        result = result.fillna(0)
        resultList.append(result)

        # 自定义排序
        p = pd.CategoricalDtype(categories=["成人湿巾", "婴儿湿巾", "儿童湿巾", "酒精湿巾", "厨房湿巾", "棉柔巾", "百亿"])
        result["品系"] = result["品系"].astype(p)
        totalResult = pd.pivot_table(result,index="品系",values=["支付金额","支付件数", "商品访客数", "成交人数"], aggfunc=[np.sum], fill_value=0)
        totalResult.columns = ['商品访客数','成交人数','支付件数','支付金额']
        totalResult = totalResult.reset_index()
        totalResult = totalResult.reindex(columns=[list(totalResult)[0], "支付金额", "支付件数", "到手价", "商品访客数", "转化率", "客单价", "成交人数", "UV价值"])
        totalOldName = list(totalResult)[0]
        totalNewName = list(totalResult)[0] + "(" + i[1] + ")"
        totalResult.rename(columns={totalOldName: totalNewName},inplace=True)

        totalResult['客单价'] = totalResult["支付金额"] / totalResult["成交人数"]
        totalResult['转化率'] = totalResult["成交人数"] / totalResult["商品访客数"]
        totalResult["到手价"] = totalResult["支付金额"] / totalResult["支付件数"]
        totalResult["UV价值"] = totalResult["支付金额"] / totalResult["商品访客数"]

        totalResult['转化率'] = totalResult["转化率"].replace(np.nan, 0)
        totalResult['转化率'] = totalResult["转化率"].apply(lambda x: format(x,'.2%'))
        totalResult['客单价'] = totalResult['客单价'].round(decimals=2)
        totalResult["UV价值"] = totalResult["UV价值"].round(decimals=2)
        totalResult["到手价"] = totalResult["到手价"].round(decimals=2)

        totalResultList.append(totalResult)

    HBData = pd.merge(resultList[0], resultList[1], on=["id", "品系"])
    TBData = pd.merge(resultList[0], resultList[2], on=["id", "品系"])

    HBDataCol = ["品系", "id", list(HBData)[2]]
    TBDataCol = ["品系", "id", list(TBData)[2]]
    for i in resultList[0].columns[3:]:
        HBData[i + "环比"] = HBData[i + "_x"] / HBData[i + "_y"] - 1
        HBData = HBData.replace([np.inf, -np.inf, np.nan], 0)
        HBData[i + "环比"] = HBData[i + "环比"].apply(lambda x: format(x,'.2%'))
        HBDataCol.append(i + "环比")

        TBData[i + "同比"] = TBData[i + "_x"] / TBData[i + "_y"] - 1
        TBData = TBData.replace([np.inf, -np.inf, np.nan], 0)
        TBData[i + "同比"] = TBData[i + "同比"].apply(lambda x: format(x,'.2%'))
        TBDataCol.append(i + "同比")

    HBData = HBData.reindex(columns=HBDataCol)
    TBData = TBData.reindex(columns=TBDataCol)
    HTBDataList = [HBData, TBData]

    writer = pd.ExcelWriter('./tmp.xlsx', mode="a")
    startCol = 0
    for i in resultList:
        i['转化率'] = i["转化率"].apply(lambda x: format(x,'.2%'))
        i.to_excel(writer, sheet_name=summaryDataSheet, startcol=startCol, index=False)
        startCol += 13
    
    startCol = 0
    for i in totalResultList:
        i.to_excel(writer, sheet_name=sumDataSheet, startcol=startCol, index=False)
        startCol += 11

    startCol = 0
    sheet_name = summaryDataSheet + "的环比&同比"
    for i in HTBDataList:
        i.to_excel(writer, sheet_name=sheet_name, startcol=startCol, index=False)
        startCol += 13

    writer.save()

def runThree(summaryDataFile=None, summaryDataSheet=None, totalDataFile=None, totalDataSheet=None, yDataFile=None, yDataSheet=None):
    summaryDataFile = summaryDataFile.replace("file:///","")
    totalDataFile = totalDataFile.replace("file:///","")
    yDataFile = yDataFile.replace("file:///","")

    # summaryDataFile = "D:\\tmp\\666.xlsx"
    # summaryDataSheet = "2"
    # totalDataFile = "D:\\tmp\\2021年维达原始数据.xlsx"
    # totalDataSheet = "2021汇总"
    # yDataFile = "D:\\tmp\\2020年维达原始数据.xlsx"
    # yDataSheet = "2020汇总"


    LDate = [1, 3, 5, 7, 8, 10, 12]
    currentCols = [0, 1, 2]
    summaryData = pd.read_excel(summaryDataFile, sheet_name=summaryDataSheet, usecols=currentCols, dtype='object')

    allTotalData = pd.read_excel(totalDataFile, sheet_name=totalDataSheet, dtype='object')

    category = summaryData.columns[0]
    dateStr = summaryData.columns[1]
    formatType = summaryData.columns[2].split(",")

    currentDate = str(datetime.datetime.now().year)
    d1 = dateStr.split("-")
    startDate = currentDate + "." + d1[0]
    endDate = currentDate + "." + d1[1]

    times = []
    cDataList = getDateList(start=startDate, end=endDate)
    times.append((cDataList, "当月"))

    if d1[0].split(".")[0] == "1":
        mStartDate = currentDate + "." + "1." + d1[0].split(".")[1]
    else:
        mStartDate = currentDate + "." + str(int(d1[0].split(".")[0]) - 1) + "." + d1[0].split(".")[1]    

    if d1[1].split(".")[0] == "1":
        mEndDate = currentDate + "." + "1." + d1[1].split(".")[1]
    else:
        mEndDate = currentDate + "." + str(int(d1[1].split(".")[0]) - 1) + "." + d1[1].split(".")[1]
    if d1[1].split(".")[1] == "31":
        if int(d1[1].split(".")[0]) - 1 not in LDate:
            mEndDate = currentDate + "." + str(int(d1[1].split(".")[0]) - 1) + "." + "30"
    mDateList = getDateList(start=mStartDate, end=mEndDate)
    times.append((mDateList, "上月"))

    yStartDate = str(int(currentDate) - 1) + "." + d1[0]
    yEndDate = str(int(currentDate) - 1) + "." + d1[1]
    yDateList = getDateList(start=yStartDate, end=yEndDate)
    times.append((yDateList, "去年"))

    resultList = []

    for i in times:

        if i[1] == "去年":
            totalDataFile = yDataFile
            totalDataSheet = yDataSheet
        allTotalData = pd.read_excel(totalDataFile, sheet_name=totalDataSheet, dtype='object')
        totalData = allTotalData[allTotalData["日期"].isin(i[0])]
        totalData = totalData[totalData["业态类型"].isin(formatType)]
        totalData = pd.pivot_table(totalData, index="品类", values=["支付金额"], aggfunc=[np.sum])
        result = pd.merge(summaryData, totalData, on="品类",how='left')
        result = result.reindex()
        result = result.reindex(columns=[result.columns[0], result.columns[-1], result.columns[1], result.columns[2]])
        result.rename(columns={result.columns[1]: "支付金额"},inplace=True)
        oldName = list(result)[2]
        newName = oldName + "(" + i[1] + ")"
        result.rename(columns={oldName: newName},inplace=True)
        
        resultList.append(result)
    
    writer = pd.ExcelWriter('./tmp.xlsx', mode="a")
    startCol = 0
    for i in resultList:
        i.to_excel(writer, sheet_name=summaryDataSheet, startcol=startCol, index=False)
        startCol += 6
    writer.save()

if __name__ == "__main__":
    runTow()