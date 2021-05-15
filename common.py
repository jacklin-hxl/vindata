
import pandas as pd

def getDateList(start, end):
    tmp = list(pd.date_range(start=start, end=end))
    dateList = [i._date_repr.replace("-","") for i in tmp]
    dateList = list(map(int, dateList))
    return dateList
