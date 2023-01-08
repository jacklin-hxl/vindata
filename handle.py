import os
import re

import numpy
import pandas

from common.tools import DateRange
from common.logger import logger


class Base:

    def __init__(self, flag=None):
        self.flag = flag

    @staticmethod
    def _read_excel(f, s, usecols=None):
        logger.debug(f"读取数据中: 路径: [{f}] sheet: [{s}] ...")
        re = pandas.read_excel(io=f, sheet_name=s, usecols=usecols, dtype="object")
        logger.debug(f"读取完成!!!")
        return re



class Sku(Base):
    RESULT_FILE_NAME = "./sku.xlsx"

    def __init__(self,
                 flag=None,
                 year=None,
                 summary_file=None,
                 summary_sheet=None,
                 source_file=None,
                 source_sheet=None):
        logger.debug("=====================>>>>> 开始处理skv")
        super().__init__(flag=flag)
        self.year = year
        self.summary_sheet = summary_sheet
        self.summary = self._read_excel(summary_file, summary_sheet)
        self.flag[0] += 10
        self.source = self._read_excel(source_file, source_sheet)
        self.flag[0] += 10
        self._range = None
        self.ans = []

    def start(self):
        """
        根据 id 获取需要算聚合数据的总个数 并调用work处理数据
        :return:
        """
        l, r = 0, 2
        times = self._amount()
        step = 60 // times
        for i in range(times):
            self.work(l, r)
            self.flag[0] += step
            l, r = l + 3, r + 3
        self.save()
        self.flag[0] += 25
        logger.debug("处理sku完成 <<<<<<<=====================")

    def work(self, l, r):
        """
        聚合数据
        :param l: id 的列数``
        :param r: 时间的 列数
        :return:
        """
        cols = self.summary.columns[l: r]
        cols = cols.copy()
        target = self.summary[cols]
        target = target.copy()
        logger.debug(f"================ 处理 [{cols[1]}] ... =====================")
        target.rename(columns={cols[0]: "id"}, inplace=True)
        _range = self._parse_date_to_range(cols, 1)
        result = self._merge(target, _range)

        self._data_handle(result)
        result = result.fillna(0)
        self.ans.append(result)
        logger.debug(f"*********************** 处理 [{cols[1]}] 完成!!! ***************************")

    def _parse_date_to_range(self, target, num):
        """
        根据开始日期/结束日期 获取日期范围
        :return:
        """
        logger.debug("获取时间范围")
        day = re.findall(r"(\d*\.\d*)日", target[num])
        start = self.year + "." + day[0]
        end = self.year + "." + day[1]
        date_range = DateRange(start, end).get_range()
        logger.debug(f"时间范围: {date_range}")
        return date_range

    def save(self):
        """
        将聚合数据写入到excel中
        :return:
        """
        if not os.path.exists(self.RESULT_FILE_NAME):
            logger.debug(f"{self.RESULT_FILE_NAME} 文件不存在，创建文件中")
            f = open(self.RESULT_FILE_NAME, "w")
            f.close()
            df = pandas.DataFrame()
            df.to_excel(self.RESULT_FILE_NAME, index=False)
        writer = pandas.ExcelWriter(self.RESULT_FILE_NAME, mode="a")
        startCol = 0
        for i in self.ans:
            logger.debug(f"转存数据到 {self.RESULT_FILE_NAME} 中...")
            i.to_excel(writer, sheet_name=self.summary_sheet, startcol=startCol, index=False)
            startCol += 12
        writer.save()
        logger.debug("转存完成!!!")

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
        logger.debug("将 时间范围内的数据 进行汇总")
        source_sum = self._source_to_sum(_range)

        target = target.dropna(axis=0, how="all")
        result = pandas.merge(target, source_sum, on="id", how='left')
        colmuns = [list(result)[0], list(result)[1], "支付金额", "支付件数", "到手价", "商品访客数", "转化率", "客单价", "成交人数", "人均购买件数",
                   "UV价值"]
        logger.debug("时间范围内的数据 汇总完成!!!")
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
        logger.debug("计算 [客单价,转化率,到手价,人均购买件数,UV价值]")
        cols[cols.columns[1]] = cols[cols.columns[1]].map(str)
        cols['客单价'] = cols["支付金额"] / cols["成交人数"]
        cols['转化率'] = cols["成交人数"] / cols["商品访客数"]
        cols["到手价"] = cols["支付金额"] / cols["支付件数"]
        cols["人均购买件数"] = cols["支付件数"] / cols["成交人数"]
        cols["UV价值"] = cols["支付金额"] / cols["商品访客数"]
        logger.debug("计算 [客单价,转化率,到手价,人均购买件数,UV价值] 完成")

    def _all_handle(self, cols):
        """
        所有商品的汇总
        :param cols:
        :return:
        """
        logger.debug("计算 所有指标数据总和 ")
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
        logger.debug("计算 所有指标数据总和 完成")

    def _format_data(self, cols):
        """
        格式化数据小数后两位
        :param cols:
        :return:
        """
        logger.debug("格式化数据")
        cols['转化率'] = cols["转化率"].apply(lambda x: format(x, '.2%'))
        cols['客单价'] = cols['客单价'].round(decimals=2)
        cols["UV价值"] = cols["UV价值"].round(decimals=2)
        cols["人均购买件数"] = cols["人均购买件数"].round(decimals=2)
        cols["到手价"] = cols["到手价"].round(decimals=2)
        cols['id'] = cols['id'].map(str)
        logger.debug("格式化数据 完成")

class Strain(Base):
    RESULT_FILE_NAME = "./strain.xlsx"

    def __init__(self,
                 is_do=True,
                 flag=None,
                 year=None,
                 summary_file=None,
                 summary_sheet=None,
                 source_file=None,
                 source_sheet=None,
                 old_file=None,
                 old_sheet=None):
        logger.debug("=====================>>>>> 开始处理品类")
        super().__init__(flag=flag)
        self.is_do = is_do
        if self.flag is None:
            self.flag = [0, False]
        self.year = year
        self.usecols = [0, 1, 2]
        self.summary_sheet = summary_sheet
        self.summary = self._read_excel(summary_file, summary_sheet, self.usecols)
        self.flag[0] += 10
        self.source = self._read_excel(source_file, source_sheet)
        self.flag[0] += 10
        self.old = self._read_excel(old_file, old_sheet)
        self.flag[0] += 10
        self._range = None
        self.ans = None
        self.amount = []
        self.total_ans = []
        self.res_list = []
        self.HTB_list = []
        self.sum_data_sheet = self.summary_sheet + "汇总"

    def start(self):
        self.prepare()
        self.flag[0] += 10
        self.work()
        self.flag[0] += 50
        self.save()
        self.flag[0] += 10
        logger.debug("处理品类完成 <<<<<=====================")
    def work(self):
        s = self.source
        for _range in self.amount:
            if _range[1] == "去年":
                s = self.old
            s_sum = self._source_to_sum(s, _range[0])
            res = self._merge(self.summary, s_sum, _range[1])
            self._single_handle(res)
            self._sort(self.res, _range[1])

        self._all_handle(self.res_list)
    def _merge(self, target, _sum, _date):
        """
        源数据和目标id的交集
        :param target:
        :param _sum:
        :return:
        """
        target = target.dropna(axis=0, how="all")
        res = pandas.merge(target, _sum, on="id", how='left')
        _col = [list(res)[0], list(res)[1], list(res)[2], "支付金额", "支付件数", "到手价", "商品访客数", "转化率", "客单价", "成交人数", "人均购买件数", "UV价值"]
        res = res.reindex(columns=_col)
        old_name = list(res)[2]
        new_name = old_name + "(" + _date + ")"
        res.rename(columns={old_name: new_name}, inplace=True)
        return res


    def _all_handle(self, res_list):
        num = 1

        if self.is_do:
            HBData = pandas.merge(res_list[0], res_list[1], on=["id", "品系"])
            HBDataCol = ["品系", "id", list(HBData)[2]]
            num = 2

        TBData = pandas.merge(res_list[0], res_list[num], on=["id", "品系"])
        TBDataCol = ["品系", "id", list(TBData)[2]]

        for i in res_list[0].columns[3:]:
            if self.is_do:
                HBData[i + "环比"] = HBData[i + "_x"] / HBData[i + "_y"] - 1
                HBData = HBData.replace([numpy.inf, -numpy.inf, numpy.nan], 0)
                HBData[i + "环比"] = HBData[i + "环比"].apply(lambda x: format(x, '.2%'))
                HBDataCol.append(i + "环比")

            TBData[i + "同比"] = TBData[i + "_x"] / TBData[i + "_y"] - 1
            TBData = TBData.replace([numpy.inf, -numpy.inf, numpy.nan], 0)
            TBData[i + "同比"] = TBData[i + "同比"].apply(lambda x: format(x, '.2%'))
            TBDataCol.append(i + "同比")

        TBData = TBData.reindex(columns=TBDataCol)
        if self.is_do:
            HBData = HBData.reindex(columns=HBDataCol)
            self.HTB_list = [HBData, TBData]
        else:
            self.HTB_list = [TBData]


    def save(self):
        if not os.path.exists(self.RESULT_FILE_NAME):
            logger.debug(f"{self.RESULT_FILE_NAME} 文件不存在，创建文件中")
            f = open(self.RESULT_FILE_NAME, "w")
            f.close()
            df = pandas.DataFrame()
            df.to_excel(self.RESULT_FILE_NAME, index=False)
        writer = pandas.ExcelWriter(self.RESULT_FILE_NAME, mode="a")

        startCol = 0
        for i in self.res_list:
            i['转化率'] = i["转化率"].apply(lambda x: format(x, '.2%'))
            i.to_excel(writer, sheet_name=self.summary_sheet, startcol=startCol, index=False)
            startCol += 13

        startCol = 0
        for i in self.total_ans:
            i.to_excel(writer, sheet_name=self.sum_data_sheet, startcol=startCol, index=False)
            startCol += 11

        startCol = 0
        if self.is_do:
            sheet_name = self.summary_sheet + "的环比&同比"
        else:
            sheet_name = self.summary_sheet + "的同比"
        for i in self.HTB_list:
            i.to_excel(writer, sheet_name=sheet_name, startcol=startCol, index=False)
            startCol += 13

        writer.save()

    def _single_handle(self, total):
        """
        计算出各个id的不同指标
        :param total:
        :return:
        """
        logger.debug("开始计算单个id的指标")
        total[total.columns[1]] = total[total.columns[1]].map(str)
        total['客单价'] = total["支付金额"] / total["成交人数"]
        total['转化率'] = total["成交人数"] / total["商品访客数"]
        total["到手价"] = total["支付金额"] / total["支付件数"]
        total["UV价值"] = total["支付金额"] / total["商品访客数"]
        total["人均购买件数"] = total["支付件数"] / total["成交人数"]
        total["品系"] = total["品系"].fillna(method='pad')

        # result['转化率'] = result["转化率"].apply(lambda x: format(x,'.2%'))
        total['转化率'] = total["转化率"].round(decimals=2)
        total['转化率'] = total["转化率"].replace(numpy.nan, 0)
        total['客单价'] = total['客单价'].round(decimals=2)
        total["UV价值"] = total["UV价值"].round(decimals=2)
        total["到手价"] = total["到手价"].round(decimals=2)
        total['id'] = total['id'].map(str)

        self.res = total.fillna(0)
        self.res_list.append(self.res)
        logger.debug("计算单个id的指标完成")

    def _sort(self, res, _date):
        logger.debug("开始排序数据")
        p = pandas.CategoricalDtype(categories=["成人湿巾", "婴儿湿巾", "儿童湿巾", "酒精湿巾", "厨房湿巾", "棉柔巾", "百亿"])
        res["品系"] = res["品系"].astype(p)
        tmp_total_res = pandas.pivot_table(res,index="品系", values=["支付金额", "支付件数", "商品访客数", "成交人数"], aggfunc=[numpy.sum], fill_value=0)
        tmp_total_res.columns = ['商品访客数', '成交人数', '支付件数', '支付金额']
        tmp_total_res = tmp_total_res.reset_index()
        tmp_total_res = tmp_total_res.reindex(columns=[list(tmp_total_res)[0], "支付金额", "支付件数", "到手价", "商品访客数", "转化率", "客单价", "成交人数", "UV价值"])
        total_old_name = list(tmp_total_res)[0]
        total_new_name = list(tmp_total_res)[0] + "(" + _date + ")"
        tmp_total_res.rename(columns={total_old_name: total_new_name}, inplace=True)
        logger.debug("排序数据完成!!!")

        logger.debug("开始计算总数据的指标之和")
        tmp_total_res['客单价'] = tmp_total_res["支付金额"] / tmp_total_res["成交人数"]
        tmp_total_res['转化率'] = tmp_total_res["成交人数"] / tmp_total_res["商品访客数"]
        tmp_total_res["到手价"] = tmp_total_res["支付金额"] / tmp_total_res["支付件数"]
        tmp_total_res["UV价值"] = tmp_total_res["支付金额"] / tmp_total_res["商品访客数"]

        tmp_total_res['转化率'] = tmp_total_res["转化率"].replace(numpy.nan, 0)
        tmp_total_res['转化率'] = tmp_total_res["转化率"].apply(lambda x: format(x,'.2%'))
        tmp_total_res['客单价'] = tmp_total_res['客单价'].round(decimals=2)
        res["人均购买件数"] = res["人均购买件数"].round(decimals=2)
        tmp_total_res["UV价值"] = tmp_total_res["UV价值"].round(decimals=2)
        tmp_total_res["到手价"] = tmp_total_res["到手价"].round(decimals=2)
        logger.debug("计算总数据的指标之和完成!!!")
        self.total_ans.append(tmp_total_res)

    def _source_to_sum(self, s, _range):
        """
        从源数据中获取指定日期内的 相同id的 商品访客数,成交人数,支付件数,支付金额
        :return:
        """
        logger.debug("开始获取指定日期内的商品指标: 商品访客数,成交人数,支付件数,支付金额")
        total = s[s["日期"].isin(_range)]
        total = pandas.pivot_table(total, index="id", values=["支付金额", "支付件数", "商品访客数", "成交人数"],
                                   aggfunc=[numpy.sum])
        total.columns = ['商品访客数', '成交人数', '支付件数', '支付金额']
        logger.debug("获取指定日期内的商品指标完成!!!")
        return total.reset_index()

    def prepare(self):
        """
        获取本月、上个月、去年时间范围
        :return:
        """

        logger.debug("获取本月、上个月、去年时间范围")
        self.summary.rename(columns={self.summary.columns[1]: "id"}, inplace=True)
        self.summary.fillna(method='pad')
        current_range = self._parse_date_to_range(self.summary.columns, 2)
        self.amount.append((current_range, "当月"))

        day = re.findall(r"(\d*\.\d*)", self.summary.columns[2])

        if self.is_do:
            logger.debug("上个月时间范围")
            old_month_start_date = self.year + "." + str(int(day[0].split(".")[0]) - 1) + "." + day[0].split(".")[1]
            old_month_end_date = self.year + "." + str(int(day[1].split(".")[0]) - 1) + "." + day[1].split(".")[1]
            old_month_range = DateRange(old_month_start_date, old_month_end_date).get_range()
            self.amount.append((old_month_range, "上月"))
            logger.debug(f"上个月范围: {old_month_range}")
        else:
            logger.debug("不求环比")

        logger.debug("去年时间范围")
        old_year_start_date = str(int(self.year) - 1) + "." + day[0]
        old_year_end_date = str(int(self.year) - 1) + "." + day[1]
        if int(self.year) % 4 == 0 and int(self.year) % 100 != 0:
            if day[1] == "2.29":
                logger("今年是闰年 且 end time 是 2.29")
                old_year_end_date = str(int(self.year) - 1) + "." + "2.28"
        old_year_range = DateRange(old_year_start_date, old_year_end_date).get_range()
        self.amount.append((old_year_range, "去年"))
        logger.debug(f"去年范围: {old_year_range}")

    def _parse_date_to_range(self, target, num):
        """
        根据开始日期/结束日期 获取日期范围
        :return:
        """
        logger.debug("获取本月时间范围")
        day = re.findall(r"(\d*\.\d*)", target[num])
        start = self.year + "." + day[0]
        end = self.year + "." + day[1]
        date_range = DateRange(start, end).get_range()
        logger.debug(f"本月时间范围: {date_range}")
        return date_range





if __name__ == '__main__':
    summary_file = "./test/数据源.xlsx"
    summary_sheet = "1"
    source_file = "./test/2022年维达原始数据.xlsx"
    source_sheet = "2022汇总"
    old_file = "./test/2021年维达原始数据(82).xlsx"
    old_sheet = "2021汇总"
    # sku = Sku('2022', summary_file, summary_sheet, source_file, source_sheet)
    # sku.start()
    strain = Strain(True, None, "2022", summary_file, summary_sheet, source_file, source_sheet, old_file, old_sheet)
    strain.start()