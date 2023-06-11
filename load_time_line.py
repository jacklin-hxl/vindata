import time
from openpyxl import Workbook
import re


open_dict = {
    "6月9号10点开抢": {"start": '2023-06-09 09:30:30', "gap": 30, "next": '2023-06-09 14:30:30'},
    "6月9号15点开抢": {"start": '2023-06-09 14:30:30', "gap": 30, "next": '2023-06-09 19:30:30'},
    "6月9号20点开抢": {"start": '2023-06-09 19:30:30', "gap": 30, "next": '2023-06-09 21:30:30'},
    "6月9号22点开抢": {"start": '2023-06-09 21:30:30', "gap": 30, "next": '2023-06-10 02:30:30'},
    "6月10号10点开抢": {"start": '2023-06-10 02:30:30', "gap": 7*60+30, "next": '2023-06-10 14:30:30'},
    "6月10号15点开抢": {"start": '2023-06-10 14:30:30', "gap": 30, "next": '2023-06-10 19:30:30'},
    "6月10号20点开抢": {"start": '2023-06-10 19:30:30', "gap": 30, "next": '2023-06-10 21:30:30'},
    "6月10号22点开抢": {"start": '2023-06-10 21:30:30', "gap": 30, "next": '2023-06-10 23:30:30'},
    "6月11号0点开抢": {"start": '2023-06-10 23:30:30', "gap": 30, "next": '2023-06-11 02:30:30'},

}

# target_date = "2023-06-09"
# target = {"10": 20, "15": 20, "20": 20}
# flag = 1
#
# time_line = []
# open_time = {}
# format_target = {}
# for i in target.keys():
#     t = target_date + " " + i + ":00:00"
#     format_target[t] = target[i]
#     tim = time.mktime(time.strptime(t, "%Y-%m-%d %H:%M:%S"))
#     tim = tim - (29*60 + 30)
#     tim = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(tim))
#     open_time[tim] = target[i]
#     open_dict = {}
#     meta = {}
#     meta["start"] = tim
#     meta["gap"] = target[i]
#     meta["next"] =

def generate_open_time(meta):
    s_stamp = convert_timestamp(meta["start"])
    # e_stamp = s_stamp + 60*30 + meta["gap"]*60
    e_stamp = convert_timestamp(meta["next"])
    continue_stamp = s_stamp + 60*30 + meta["gap"]*60
    all_time = []

    all_time.append(s_stamp)
    s_stamp = s_stamp + 60*30
    all_time.append(s_stamp)

    re, con = do_gap_handle(all_time, s_stamp, e_stamp, continue_stamp)
    meta["continue"] = all_time[con:-1]
    meta["open"] = all_time[:con]

def do_gap_handle(all_time, s, e, c):
    do_gap = [60, 4 * 60, 5 * 60, 10 * 60, 10 * 60, 10 * 60, 10 * 60, 10 * 60]
    con = ""
    while s < e:
        for i in do_gap:
            s += i
            if s >= e:
                all_time.append(s)
                break
            if s >= c > (s - i):
                con = len(all_time)
            all_time.append(s)
    return all_time, con


def is_int_point(s):
    s = s.split(" ")[1].split(":")[1]
    if s == "00":
        return True
    else:
        return False

def convert_timestamp(s):
    return time.mktime(time.strptime(s, "%Y-%m-%d %H:%M:%S"))

def convert_strtime(t):
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(t))


def format_time(time_line):
    format_time_line = []
    for t in time_line:
        format_time_line.append(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(t)))
    return format_time_line


def generate_continue_time(s, e):
    continue_list = []
    c_s = convert_timestamp(s)
    c_e = convert_timestamp(e)
    continue_list.append(c_s)
    do_list = [60, 4 * 60, 5 * 60, 10 * 60, 10 * 60, 10 * 60, 10 * 60, 10 * 60, 10 * 60]
    if is_int_point(s):
        while c_s < c_e:
            for i in do_list:
                c_s = c_s + i
                continue_list.append(c_s)
    else:
        while is_int_point(convert_strtime(c_s)) is False and c_s < c_e:
            c_s += 10*60
            continue_list.append(c_s)
        while c_s < c_e:
            for i in do_list:
                c_s = c_s + i
                if is_int_point(convert_strtime(c_s)):
                    continue_list.append(c_s)
                    break
                continue_list.append(c_s)
    return continue_list

res = []
for k, v in list(open_dict.items()):
    generate_open_time(open_dict[k])

for k,v in open_dict.items():
    continue_description = k.replace("开抢", "结束")
    target = re.findall("月(.*)号", k)[0]
    new = str(int(target) + 1)
    tg = re.findall("号(.*)点", k)[0]
    tg = f"号{tg}点"
    to_tg = "号0点"
    continue_description = continue_description.replace(target, new)
    continue_description = continue_description.replace(tg, to_tg)
    _open = format_time(open_dict[k]["open"])
    for i in range(len(_open)):
        tmp = []
        tmp.append(_open[i])
        tmp.append(k)
        res.append(tmp)
    _continue = format_time(open_dict[k]["continue"])
    for j in range(len(_continue)):
        tmp = []
        tmp.append(_continue[j])
        tmp.append(continue_description)
        res.append(tmp)

workbook = Workbook()
sheet = workbook.active
for row in res:
    sheet.append(row)
workbook.save("./tmp.xlsx")
print("s")

