#!/usr/bin/python3

import json
import sys
from configparser import ConfigParser
from datetime import datetime
from datetime import timedelta
from pprint import pprint

from dateutil.relativedelta import relativedelta
from pypinyin import lazy_pinyin

from lib.demail import Email
from lib.tools import Tools
from smtplib import SMTPRecipientsRefused

SEND_EMAIL = False
TARGET = None # 这个要设置为将要统计的月份(如：2016-6)，如果为None则从文件读取
CONFIG_FILE = "config.ini"

TEMPLATE = """
{name}同学：
<p>
您好！<br>
{year}年{month}月的加班调休时长记录已产生，记录了您{year}年{month}月1日-{year}年{month}月{end_day}日的变动信息，现为您诚意奉上，仅供您参考。<br>
截止至{old_month}月{old_end_day}日24时，您剩余可调休时长{last_remaining}小时；<br>
{month}月份新增{overtime}小时，已调休使用{used_overtime}小时，深之度支付加班费时长{paid}小时；<br>
截止至{month}月{end_day}日24时，结余{remaining}小时。<br>
感谢您长久以来的辛勤付出，祝您工作顺心！<br>
</p>
"""

USER_FILTER = []


class PaidLeave:

    def __init__(self, target):
        self.target = target
        self.tools = Tools()
        if SEND_EMAIL:
            self.email = Email()


    def get_overtime(self):
        """
        当月新增
        """
        data = {} # {user: time}

        month_str = self.tools.get_month_str(self.target)
        filename = "data/%s/overtime.xlsx" % month_str
        _data = self.tools.get_excel_data(filename, ["姓名", "起始时间", "结束时间"], 1)
        i = 0
        for name in _data.get("姓名"):
            start_time_str = _data.get("起始时间")[i]
            end_time_str = _data.get("结束时间")[i]
            start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M")
            end_time = datetime.strptime(end_time_str, "%Y-%m-%d %H:%M")
            del_time = round((end_time - start_time).seconds / 3600, 1)
            # someone may has multiple overtime record, accumulate those data
            _time = data.get(name, 0)
            data[name] = _time + del_time
            i += 1
        return data


    def get_last_remaining(self):
        """
        上月剩余
        """
        data = {} # {user: time}

        month_str = self.tools.get_month_str(self.tools.get_last_month_dt(self.target))
        filename = "data/%s/all.xlsx" % month_str
        _data = self.tools.get_excel_data(filename, ["姓名", "剩余可用"])
        i = 0
        for name in _data.get("姓名"):
            _time = data.get(name, 0)
            data[name] = _time + _data.get("剩余可用")[i]
            i += 1

        return data


    def get_used_overtime(self):
        """
        当月已用
        """
        data = {} # {user: time}
        month_str = self.tools.get_month_str(self.target)
        filename = "data/%s/leave.xlsx" % month_str
        excel_data = self.tools.get_excel_data(filename, ["发起人姓名", "请假天数", "请假类型", "审批结果"])
        i = 0
        for name in excel_data.get("发起人姓名"):
            if "倒休" in excel_data.get("请假类型")[i] and "同意" in excel_data.get("审批结果")[i]:
                time = data.get(name, 0)
                data[name] = time + 8 * float(excel_data.get("请假天数")[i])
            i += 1
        return data


    def get_paid(self):
        """
        当月支付
        """
        data = {} # {user: time}
        month_str = self.tools.get_month_str(self.target)
        filename = "data/%s/overtime.xlsx" % month_str
        excel_data = self.tools.get_excel_data(filename, ["姓名", "起始时间", "结束时间", "是否支付"], 1)
        i = 0

        if not len(excel_data):
            return {}

        for name in excel_data.get("姓名"):
            if excel_data.get("是否支付")[i] == 1:
                start_time_str = excel_data.get("起始时间")[i]
                end_time_str = excel_data.get("结束时间")[i]
                start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M")
                end_time = datetime.strptime(end_time_str, "%Y-%m-%d %H:%M")
                del_time = round((end_time - start_time).seconds / 3600, 1)
                # someone may has multiple overtime record, accumulate those data
                _time = data.get(name, 0)
                data[name] = _time + del_time
            i += 1
        return data


    def get_all_users_data(self):
        users = {}
        with open("data/user.json") as fp:
            users = json.load(fp)

        return users


    def work(self):
        user_data = self.get_all_users_data()
        last_remaining = self.get_last_remaining()
        paid = self.get_paid()
        overtime = self.get_overtime()
        used_overtime = self.get_used_overtime()

        for user_guid, user in user_data.items():
            name = user.get("name")
            user["paid"] = paid.get(name, 0)
            user["overtime"] = overtime.get(name, 0)
            user["used_overtime"] = used_overtime.get(name, 0)
            user["last_remaining"] = last_remaining.get(name, 0)
            user["remaining"] = user["last_remaining"] + user["overtime"] - user["used_overtime"]  - user["paid"]

        self.generate_excel(user_data)


    def generate_excel(self, user_data):
        first_date = self.target.replace(day=1)
        last_end_date = first_date - timedelta(days=1)
        last_end_date_str = "%s月%s日" % (last_end_date.month, last_end_date.day)
        excel_data = [["姓名", "截止%s剩余" % last_end_date_str, "%s月份新增" % self.target.month, "%s月份已用" % self.target.month, "%s月份支付" % self.target.month, "剩余可用"], ]
        TTT = False
        index = 0
        user_data_list = user_data.values()
        user_data_list = sorted(user_data_list, key=name_sortor)
        for user in user_data_list:
            if user.get("paid") == 0 and user.get("overtime") == 0 \
                    and user.get("used_overtime") == 0 and user.get("last_remaining") == 0 \
                    and user.get("remaining") == 0:
                continue
            end_day = last_end_date.day
            last_month_date = self.tools.get_last_month_dt(self.target)
            old_end_day = last_end_date.day

            after_month = self.target + relativedelta(months=1)
            end_date = after_month - timedelta(days=1)

            params = {
                    "name": user.get("name"), 
                    "last_remaining": user.get("last_remaining"),
                    "overtime": user.get("overtime"),
                    "used_overtime": user.get("used_overtime"),
                    "paid": user.get("paid"),
                    "remaining": user.get("remaining"),
                    "year": self.target.year,
                    "month": self.target.month,
                    "end_day": end_date.day,
                    "old_month": last_month_date.month,
                    "old_end_day": old_end_day
                    }
            content = TEMPLATE.format(**params)
            print("==========================================================")
            print(content)
            if SEND_EMAIL:
                print("----------------------------------------------------------")
                receiver = user.get("email").replace("linuxdeepin.com", "deepin.com")
                print("sending email to: ", receiver)

                try:
                    if user.get("name") in USER_FILTER:
                        print("user is in the filter list, skip sending...")
                    else:
                        subject = "%s年%s月调休统计" % (self.target.year, self.target.month)
                        self.email.send(receiver, subject, content)

                except SMTPRecipientsRefused as e:
                    print("failed to sending email")
                    print(e)

                except Exception as e:
                    print("failed to sending email")
                    print(e)

                print("finish.")

            print("==========================================================")

            col = [user.get("name"), user.get("last_remaining"), user.get("overtime"), user.get("used_overtime"), user.get("paid"), user.get("remaining")]
            excel_data.append(col)
            index += 1

        self.tools.write_to_execl("data/%s/all.xlsx" % self.tools.get_month_str(self.target), excel_data)


def name_sortor(item):
    return "".join(lazy_pinyin(item.get("name"), errors="ignore"))


def month_str_to_date(month_str):
    year = month_str.split("-")[0]
    month = month_str.split("-")[1]
    d = datetime(int(year), int(month), 1)
    return d


def get_month_from_config():
    c = ConfigParser()
    c.read(CONFIG_FILE)
    month_str = c["default"]["month"]
    return month_str


if __name__ == "__main__":
    if "--mail" in sys.argv:
        SEND_EMAIL = True

    target = TARGET if TARGET else get_month_from_config()
    d = month_str_to_date(target)
    pl = PaidLeave(d)
    pl.work()
