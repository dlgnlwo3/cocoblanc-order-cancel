if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from selenium import webdriver
from dtos.gui_dto import GUIDto
from dtos.product_dto import ProductDto

from PySide6.QtCore import SignalInstance

from common.chrome import get_chrome_driver, get_chrome_driver_new
from common.selenium_activities import close_new_tabs, alert_ok_try, wait_loading, send_keys_to_driver
from common.account_file import AccountFile

from api.eleven_street_api import ElevenStreetAPI

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver import ActionChains

import pandas as pd
import asyncio
import pyperclip
import time
import re

from datetime import datetime, timedelta
from collections import defaultdict

from process.ezadmin import Ezadmin
from process.kakaotalk_store import KakaoTalkStore
from process.wemakeprice import Wemakeprice
from process.zigzag import Zigzag
from process.bflow import Bflow
from process.coupang import Coupang
from process.eleven_street import ElevenStreet
from process.ticketmonster import TicketMonster
from process.naver import Naver


class CocoblancOrderCancelProcess:
    def __init__(self):
        self.driver: webdriver.Chrome = get_chrome_driver(is_headless=False, is_secret=False)
        self.default_wait = 10
        self.driver.implicitly_wait(self.default_wait)
        try:
            self.driver.maximize_window()
        except Exception as e:
            print(str(e))

    def setGuiDto(self, guiDto: GUIDto):
        self.guiDto = guiDto

    def setLogger(self, log_msg):
        self.log_msg: SignalInstance = log_msg

    def get_dict_account(self):
        df_accounts = AccountFile(self.guiDto.account_file).df_account
        df_accounts = df_accounts.fillna("")
        dict_accounts = {}
        for index, row in df_accounts.iterrows():
            channel = str(row["채널명"])
            domain = str(row["도메인"])
            account_id = str(row["ID"])
            account_pw = str(row["PW"])
            url = str(row["URL"])
            api_key = str(row["API_KEY"])
            dict_accounts[channel] = {"도메인": domain, "ID": account_id, "PW": account_pw, "URL": url, "API_KEY": api_key}
        return dict_accounts

    # 전체작업 시작
    def work_start(self):
        print(f"CocoblancOrderCancelProcess: work_start")

        try:
            self.dict_accounts = self.get_dict_account()

            # 쇼핑몰의 수 만큼 작업 (계정 엑셀 파일의 행)
            for account in self.dict_accounts:
                try:
                    dict_account = self.dict_accounts[account]

                    if account == "이지어드민":
                        ezadmin = Ezadmin(self.log_msg, self.driver, dict_account)
                        ezadmin.login()
                        self.cs_screen_tab = ezadmin.switch_to_cs_screen()

                    # if account == "카카오톡스토어":
                    #     kakaotalk_store = KakaoTalkStore(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     kakaotalk_store.work_start()

                    # if account == "위메프":
                    #     wemakeprice = Wemakeprice(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     wemakeprice.work_start()

                    # if account == "티몬":
                    #     ticketmonster = TicketMonster(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     ticketmonster.work_start()

                    # if account == "지그재그":
                    #     zigzag = Zigzag(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     zigzag.work_start()

                    # if account == "브리치":
                    #     bflow = Bflow(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     bflow.work_start()

                    # if account == "쿠팡":
                    #     coupang = Coupang(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     coupang.work_start()

                    # if account == "11번가":
                    #     eleven_street = ElevenStreet(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     eleven_street.work_start()

                    if account == "네이버":
                        naver = Naver(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                        naver.work_start()

                except Exception as e:
                    print(str(e))

        except Exception as e:
            print(str(e))


if __name__ == "__main__":
    # process = CocoblancOrderCancelProcess()
    # process.work_start()
    pass
