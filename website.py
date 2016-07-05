#!/usr/bin/python3

import argparse
import json
import operator
import os
import re
import requests
import time
import sys

from datetime import datetime
from datetime import timedelta
from configparser import ConfigParser

# third lib
import requests
import xlsxwriter

from pyvirtualdisplay import Display
from selenium import webdriver

# my lib
from lib.demail import Email


INPUT = "user.json"
OUTPUT = "data/user.json"

SEND_DAY = 1  # day of every month
USER_CONF_PATH = "%s/.AutoScriptConfig/tower-overtime-reportor/user.ini" % os.getenv("HOME")
MEMBER_URL = "https://tower.im/teams/35e3a49a6e2e40fa919070f0cd9706c8/members/"

BASE_TOWER_URL = "tower.im/api/v2"
TOWER_API = "https://%s" % BASE_TOWER_URL

GROUP_FILTER_LIST = ["79d4b0d56cac41bd931fc772365772c4", # 总经办
        "a8b36416233c441da3f0a69ae0d4b1d9", # 江南援兵
        "7aa50b3fa0ca4576b7731b5fb798cd4b", # 合作伙伴
        "7c93c3b566584e2cbc2a7abed47282b1", # 网易合作
        "41f6be08aa42424ab29d02d99b65c9d1", # 机器人团队
        ] 

EMAIL_FILTER_LIST = []


class ConfigController:

    def __init__(self):
        self.tower_token = ""

    def get_login_info(self):
        config = ConfigParser()
        config.read(USER_CONF_PATH)
        username = config["USER"]["UserName"]
        passwd = config["USER"]["UserPWD"]
        return username, passwd


    def get_tower_token(self):
        if self.tower_token == "":
            config = ConfigParser()
            config.read(USER_CONF_PATH)
            username = config.get("USER", "UserName")
            passwd = config.get("USER", "UserPWD")
            client_id = config.get("DEEPIN", "ClientId")
            client_secret = config.get("DEEPIN", "ClientSecret")

            url = "https://%s:%s@%s/oauth/token" % (client_id, client_secret, BASE_TOWER_URL)
            d = {"grant_type":"password", "username": username, "password": passwd}
            success, data = self.__sendRequest(url, d)

            if success:
                self.tower_token = data.get("access_token")
            else:
                print("E: get tower access token error", file=sys.stderr)

        return self.tower_token


    def __sendRequest(self, url, d={}, h={}, method='POST'):
        if method == 'POST':
            resp = requests.post(url, data=d, headers=h)
        elif method == 'GET':
            resp = requests.get(url, data=d, headers=h)
        else:
            print("request method not supported")
            return False, None

        if resp.ok:
            return True, resp.json()

        print ("E: send request error: %s" % resp.text, file=sys.stderr)
        return False, None


class BrowserController:

    def __init__(self):
        self.browser = webdriver.Firefox()
        self.cc = ConfigController()
        (username, passwd) = self.cc.get_login_info()
        self.login(username, passwd)


    def login(self, username, passwd):

        print("login to tower ...")
        login_url = "https://tower.im/users/sign_in"
        self.browser.get(login_url)
        unEL = self.browser.find_element_by_id("email")
        pwdEL = self.browser.find_element_by_name("password")
        unEL.send_keys(username)
        pwdEL.send_keys(passwd)
        unEL.submit()

        time.sleep(5)

        # check login status
        if self.browser.current_url == "https://tower.im/teams/35e3a49a6e2e40fa919070f0cd9706c8/projects/":
            print ("login successfully")
            return True

        else:
            print ("login error, current url (%s) does not match." % self.browser.current_url)
            return False


    def get_user_info(self, existed_user_data):
        data = {}

        self.browser.get(MEMBER_URL)
        groupListEL = self.browser.find_element_by_class_name("grouplists")
        groupELs = groupListEL.find_elements_by_class_name("group")

        print("collecting all user guid...")
        for g in groupELs:
            groupGuid = g.get_attribute("data-guid").strip()

            # filter the group
            if groupGuid in GROUP_FILTER_LIST:
                continue

            member_els = g.find_elements_by_class_name("member")
            for member_el in member_els:
                guid = member_el.get_attribute("data-guid").strip()
                name_el = member_el.find_elements_by_class_name("name")[0]
                name = name_el.text.strip()
                d = {"name":name}
                data[guid] = d

        for guid, user_info in data.items():
            print("getting %s email..." % user_info.get("name"))
            if guid in existed_user_data:
                print("%s existed, skip getting from browser" % user_info.get("name"))
                email = existed_user_data.get(guid).get("email")
            else:
                email = self.get_user_email(guid)

            if email in EMAIL_FILTER_LIST:
                continue

            user_info["email"] = email

        return data


    def get_user_email(self, guid):
        url = "https://tower.im/members/%s/" % guid
        self.browser.get(url)

        # filter users
        emailEL = self.browser.find_element_by_class_name("email")
        email = emailEL.text.strip()
        return email


class OvertimeAnalyze:

    def __init__(self):
        self.cc = ConfigController()

    def work(self):
        self.prepare_user_info()

    def prepare_user_info(self):
        user_info = self.get_user_info(self.existed_user_data())
        with open(OUTPUT, "w") as fp:
            json.dump(user_info, fp)

    def existed_user_data(self):
        data = {}
        if os.path.exists(INPUT):
            with open(INPUT) as fp:
                data = json.load(fp)
        return data

    
    def get_user_info(self, existed_user_data):
        bc = BrowserController()
        user_info = bc.get_user_info(existed_user_data)
        return user_info


if __name__ == "__main__":
    display = Display(visible=0, size=(1366, 768))
    display.start()
    oa = OvertimeAnalyze()
    oa.work()
    display.stop()
