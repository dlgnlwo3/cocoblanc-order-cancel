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

    def get_naver_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
        try:
            driver.get("https://admin.pay.naver.com/o/v3/claim/cancel?summaryInfoType=CANCEL_REQUEST_C1")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "취소관리")]'))
            )
            time.sleep(0.5)

            # 주문번호 목록 -> 네이버의 경우 구매자 연락처
            # $x('//tr/td[@data-column-name="receiverTelNo1"]/div')
            order_number_list = driver.find_elements(By.XPATH, '//tr/td[@data-column-name="receiverTelNo1"]/div')
            phone_number_pattern = r"^01[016789]-\d{3,4}-\d{4}$"
            virtual_number_pattern = r"050\d{1}-\d{3,4}-\d{3,4}"
            for order_number in order_number_list:
                order_number = order_number.get_attribute("textContent")
                if re.search(phone_number_pattern, order_number) or re.search(virtual_number_pattern, order_number):
                    order_cancel_list.append(order_number)
                else:
                    print(f"{order_number}는 전화번호, 안심번호 양식이 아닙니다.")

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

        return order_cancel_list

    def naver_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

        try:
            driver.get("https://admin.pay.naver.com/o/v3/claim/cancel?summaryInfoType=CANCEL_REQUEST_C1")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "취소관리")]'))
            )
            time.sleep(0.5)

            # 취소 품목의 행 번호 -> rside
            # $x('//tr[./td[@data-column-name="receiverTelNo1"]/div[text()="010-4134-0899"]]')
            order_cancel_target_row = driver.find_element(
                By.XPATH, f'//tr[./td[@data-column-name="receiverTelNo1"]/div[text()="{order_cancel_number}"]]'
            ).get_attribute("data-row-key")

            # 취소 품목 행과 동일한 lside에 존재하는 radiobutton
            # $x('//div[@class="tui-grid-lside-area"]//tr[@data-row-key="1"]//input[contains(@type, "radio")]')
            order_cancel_target_checkbox = driver.find_element(
                By.XPATH,
                f'//div[@class="tui-grid-lside-area"]//tr[@data-row-key="{order_cancel_target_row}"]//input[contains(@type, "radio")]',
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(0.5)

            # 취소 완료처리 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[./span[text()="취소 완료처리"]]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # 새 창 열림 or alert ['선택된 상품주문건이 없습니다.', '취소 승인 처리가 불가능한 상태입니다. 클레임 처리상태를 확인해 주세요.']
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            naver_order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(naver_order_cancel_tab)
                time.sleep(1)

                # 취소비용 청구관련 구매자에게 전하실 말씀 (최대 500자) [구매자에게 전하실 말씀을 입력하십시오.]
                seller_memo = "전체주문이 취소되어 클레임 비용을 청구하지 않습니다."
                input_sellerMemoByCancel = driver.find_element(By.XPATH, '//input[@id="sellerMemoByCancel"]')
                try:
                    input_sellerMemoByCancel.clear()
                    input_sellerMemoByCancel.send_keys(seller_memo)
                except Exception as e:
                    print(str(e))
                finally:
                    time.sleep(1)

                naver_order_cancel_button = driver.find_element(By.XPATH, '//a[./span[text()="저장"]]')
                driver.execute_script("arguments[0].click();", naver_order_cancel_button)
                time.sleep(1)

                # 클레임 비용이 0원인 건은 승인처리시 즉시 환불이 시도됩니다. 승인처리 하시겠습니까? alert
                alert_msg = ""
                try:
                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    alert = driver.switch_to.alert
                    alert_msg = alert.text
                except Exception as e:
                    print(f"no alert")

                print(f"{alert_msg}")

                if "승인처리 하시겠습니까" in alert_msg:
                    alert.accept()
                    time.sleep(1)

                    # 정상적으로 저장되었습니다. alert
                    try:
                        WebDriverWait(driver, 10).until(EC.alert_is_present())
                    except Exception as e:
                        print(f"no alert")
                        raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

                    alert_ok_try(driver)
                    time.sleep(0.5)

                elif alert_msg != "":
                    raise Exception(f"{account} {order_cancel_number}: {alert_msg}")

                else:
                    raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))
                driver.close()
                if account in str(e):
                    raise Exception(str(e))

            finally:
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{account} {order_cancel_number}: 해당 주문이 존재하지 않습니다.")

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

                    if account == "지그재그":
                        zigzag = Zigzag(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                        zigzag.work_start()

                    if account == "브리치":
                        bflow = Bflow(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                        bflow.work_start()

                    if account == "쿠팡":
                        coupang = Coupang(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                        coupang.work_start()

                    if account == "11번가":
                        eleven_street = ElevenStreet(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                        eleven_street.work_start()

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
