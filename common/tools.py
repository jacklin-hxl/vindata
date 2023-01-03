import datetime

import pandas as pd


def getDateList(start, end):
    tmp = list(pd.date_range(start=start, end=end))
    dateList = [i._date_repr.replace("-","") for i in tmp]
    dateList = list(map(int, dateList))
    return dateList


class DateRange:

    def __init__(self, start, end):
        self.start = datetime.datetime.strptime(start, "%Y.%m.%d").strftime("%Y%m%d")
        self.end = datetime.datetime.strptime(end, "%Y.%m.%d").strftime("%Y%m%d")

    def get_range(self):
        tmp = list(pd.date_range(start=self.start, end=self.end))
        dateList = [i._date_repr.replace("-", "") for i in tmp]
        dateList = list(map(int, dateList))
        return dateList

    def __generator(self):
        cur = self.start
        dt = datetime.datetime.strptime(self.start, "%Y%m%d")
        while True:
            if cur > self.end:
                raise StopIteration
            yield int(dt.strftime("%Y%m%d"))
            dt = dt + datetime.timedelta(days=1)
            cur = dt.strftime("%Y.%m.%d")


if __name__ == '__main__':

    # print([i for i in get_date_range("20220801", "20220910")])
    a = DateRange("2022.12.22", "2022.12.23")
    print(getDateList("2022.12.22", "2022.12.23"))
