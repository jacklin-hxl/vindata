import os
import re

import numpy
import pandas

from common.tools import DateRange


class Base:

    @staticmethod
    def _read_excel(f, s):
        print(f"读取数据中: 路径: [{f}] sheet: [{s}] ...")
        re = pandas.read_excel(io=f, sheet_name=s, dtype="object")
        print(f"读取完成!!!")
        return re

class Sku(Base):
    RESULT_FILE_NAME = "./sku.xlsx"

    def __init__(self,
                 year=None,
                 summary_file=None,
                 summary_sheet=None,
                 source_file=None,
                 source_sheet=None):

        self.year = year
        self.summary_sheet = summary_sheet
        self.summary = self._read_excel(summary_file, summary_sheet)
        self.source = self._read_excel(source_file, source_sheet)
        self._range = None
        self.ans = []

    def start(self):
        """
        根据 id 获取需要算聚合数据的总个数 并调用work处理数据
        :return:
        """
        l, r = 0, 2
        times = self._amount()
        for i in range(times):
            self.work(l, r)
            l, r = l + 3, r + 3
        self.save()

    def work(self, l, r):
        """
        聚合数据
        :param l: id 的列数
        :param r: 时间的 列数
        :return:
        """
        cols = self.summary.columns[l: r]
        cols = cols.copy()
        target = self.summary[cols]
        target = target.copy()
        print(f"================ 处理 [{cols[1]}] ... =====================")
        target.rename(columns={cols[0]: "id"}, inplace=True)
        _range = self._parse_date_to_range(cols)
        result = self._merge(target, _range)

        self._data_handle(result)
        result = result.fillna(0)
        self.ans.append(result)
        print(f"*********************** 处理 [{cols[1]}] 完成!!! ***************************")

    def save(self):
        """
        将聚合数据写入到excel中
        :return:
        """
        if not os.path.exists(self.RESULT_FILE_NAME):
            print(f"{self.RESULT_FILE_NAME} 文件不存在，创建文件中")
            f = open(self.RESULT_FILE_NAME, "w")
            f.close()
            df = pandas.DataFrame()
            df.to_excel(self.RESULT_FILE_NAME, index=False)
        writer = pandas.ExcelWriter(self.RESULT_FILE_NAME, mode="a")
        startCol = 0
        for i in self.ans:
            print(f"转存数据到 {self.RESULT_FILE_NAME} 中...")
            i.to_excel(writer, sheet_name=self.summary_sheet, startcol=startCol, index=False)
            startCol += 12
        writer.save()

    def _amount(self):
        amount = 0
        for i in self.summary.columns:
            if "id" in i:
                amount += 1
        return amount

    def _merge(self, target, _range):
        """
        取源数据和目标id的交集，即获取目标id
        "支付金额","支付件数", "商品访客数", "成交人数" 之和
        :param target: 目标id
        :param source_sum: 源数据指定日期的不同id的相关数据之和
        :return:
        """
        print("将 时间范围内的数据 进行汇总")
        source_sum = self._source_to_sum(_range)

        target = target.dropna(axis=0, how="all")
        result = pandas.merge(target, source_sum, on="id", how='left')
        colmuns = [list(result)[0], list(result)[1], "支付金额", "支付件数", "到手价", "商品访客数", "转化率", "客单价", "成交人数", "人均购买件数",
                   "UV价值"]
        print("时间范围内的数据 汇总完成!!!")
        return result.reindex(columns=colmuns)

    def _source_to_sum(self, range):
        """
        从源数据中获取指定日期内的 相同id的 ‘商品访客数', '成交人数', '支付件数','支付金额’ 总和
        :param range: 日期范围
        :return:
        """
        total = self.source[self.source["日期"].isin(range)]
        total = pandas.pivot_table(total, index="id", values=["支付金额", "支付件数", "商品访客数", "成交人数"], aggfunc=[numpy.sum])
        total.columns = ['商品访客数', '成交人数', '支付件数', '支付金额']
        return total.reset_index()

    def _parse_date_to_range(self, target):
        """
        根据开始日期/结束日期 获取日期范围
        :return:
        """
        print("获取时间范围")
        day = re.findall(r"(\d*\.\d*)日", target[1])
        start = self.year + "." + day[0]
        end = self.year + "." + day[1]
        date_range = DateRange(start, end).get_range()
        print(f"时间范围: {date_range}")
        return date_range

    def _data_handle(self, cols):
        self._single_handle(cols)
        self._all_handle(cols)
        self._format_data(cols)

    def _single_handle(self, cols):
        """
        计算出单品不同指标
        :param cols:
        :return:
        """
        print("计算 [客单价,转化率,到手价,人均购买件数,UV价值]")
        cols[cols.columns[1]] = cols[cols.columns[1]].map(str)
        cols['客单价'] = cols["支付金额"] / cols["成交人数"]
        cols['转化率'] = cols["成交人数"] / cols["商品访客数"]
        cols["到手价"] = cols["支付金额"] / cols["支付件数"]
        cols["人均购买件数"] = cols["支付件数"] / cols["成交人数"]
        cols["UV价值"] = cols["支付金额"] / cols["商品访客数"]
        print("计算 [客单价,转化率,到手价,人均购买件数,UV价值] 完成")

    def _all_handle(self, cols):
        """
        所有商品的汇总
        :param cols:
        :return:
        """
        print("计算 所有指标数据总和 ")
        resultSum = cols.sum()
        # 支付金额
        totalAmount = resultSum[2]

        # 支付件数
        totalNumber = resultSum[3]

        # 到手价
        averagePrice = totalAmount / totalNumber

        # 商品访客数
        visitors = resultSum[5]

        # 转化率
        conversionRate = resultSum[-3] / resultSum[5]

        # 客单价
        unitPrice = resultSum[2] / resultSum[-3]

        # 成交人数
        payNumber = resultSum[-3]

        # 人均购买件数
        RJNumber = totalNumber / payNumber

        # UV价值
        UVWorth = totalAmount / visitors

        summaryValue = ["汇总", "", totalAmount, totalNumber, averagePrice, visitors, conversionRate, unitPrice,
                        payNumber, RJNumber, UVWorth]
        cols.loc[cols.shape[0]] = dict(zip(cols.columns, summaryValue))
        print("计算 所有指标数据总和 完成")

    def _format_data(self, cols):
        """
        格式化数据小数后两位
        :param cols:
        :return:
        """
        print("格式化数据")
        cols['转化率'] = cols["转化率"].apply(lambda x: format(x, '.2%'))
        cols['客单价'] = cols['客单价'].round(decimals=2)
        cols["UV价值"] = cols["UV价值"].round(decimals=2)
        cols["人均购买件数"] = cols["人均购买件数"].round(decimals=2)
        cols["到手价"] = cols["到手价"].round(decimals=2)
        cols['id'] = cols['id'].map(str)
        print("格式化数据 完成")

if __name__ == '__main__':
    summary_file = "./test/数据源.xlsx"
    summary_sheet = "4"
    source_file = "./test/2022年维达原始数据.xlsx"
    source_sheet = "2022汇总"
    sku = Sku('2022', summary_file, summary_sheet, source_file, source_sheet)
    sku.start()
