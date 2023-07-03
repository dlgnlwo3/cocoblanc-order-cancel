if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from selenium import webdriver
from dtos.product_dto import ProductDto

from PySide6.QtCore import SignalInstance

from process.ezadmin import check_order_cancel_number_from_ezadmin

from common.chrome import get_chrome_driver, get_chrome_driver_new
from common.selenium_activities import close_new_tabs, alert_ok_try, wait_loading, send_keys_to_driver

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


class ElevenStreet:
    def __init__(self, log_msg, driver, cs_screen_tab, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.cs_screen_tab: str = cs_screen_tab
        self.dict_account: dict = dict_account

        self.shop_name = "11번가"
        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]
        self.api_key = self.dict_account["API_KEY"]

        self.default_wait = 10

    def login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 바로 스토어 선택 화면으로 이동합니다.
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '//h1[contains(text(), "셀러오피스")]'))
            )
            time.sleep(2)

        except Exception as e:
            pass

        try:
            driver.implicitly_wait(1)

            id_input = driver.find_element(By.XPATH, '//input[@id="loginName"]')
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(self.login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@id="passWord"]')
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(self.login_pw)

            login_button = driver.find_element(By.XPATH, '//button[@value="로그인"]')
            time.sleep(0.2)
            login_button.click()
            time.sleep(0.2)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            main_page_link = driver.find_element(By.XPATH, '//h1[./a[contains(text(), "Seller Office")]]')

        except Exception as e:
            self.log_msg.emit(f"{self.shop_name} 로그인 실패")
            print(e)
            raise Exception(f"{self.shop_name} 로그인 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

    def get_order_list(self):
        driver = self.driver

        order_list = []

        try:
            APIBot = ElevenStreetAPI(self.api_key)

            driver.get("https://soffice.11st.co.kr/view/6209?preViewCode=D")

            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//iframe[@title="취소관리"]')))
            time.sleep(0.5)

            # 조회 기간은 최대 30일 YYYYMMDDhhmm  'strftime("%Y%m%d%H%M")' 활용
            now = datetime.now()
            startTime = str((now - timedelta(days=30)).strftime("%Y%m%d")) + "0000"
            endTime = str(now.strftime("%Y%m%d")) + "2359"

            # 취소처리 API에는 ordPrdCnSeq, ordNo, ordPrdSeq (클레임번호, 주문번호, 주문순번) 총 세가지 정보가 필요함
            cancelorders = asyncio.run(APIBot.get_cancelorders_from_date(startTime, endTime))

            claim_data = []
            try:
                if type(cancelorders["ns2:orders"]["ns2:order"]) == dict:
                    api_cancelorder_list = [cancelorders["ns2:orders"]["ns2:order"]]
                else:
                    api_cancelorder_list = cancelorders["ns2:orders"]["ns2:order"]

                for api_cancelorder in api_cancelorder_list:
                    product_dto = ProductDto()

                    # ordNo -> 주문번호
                    order_number = api_cancelorder["ordNo"]
                    product_dto.order_number = order_number

                    # ordPrdSeq -> 주문상세번호
                    order_detail_number = api_cancelorder["ordPrdSeq"]
                    product_dto.order_detail_number = order_detail_number

                    # slctPrdOptNm -> 판매처 옵션
                    product_option = api_cancelorder["slctPrdOptNm"]
                    product_dto.product_option = product_option

                    # ordCnQty -> 클레임 수량
                    product_qty = api_cancelorder["ordCnQty"]
                    product_dto.product_qty = product_qty

                    # ordPrdCnSeq -> 외부몰 클레임 번호 (취소작업에 반드시 필요한 정보)
                    product_name = api_cancelorder["ordPrdCnSeq"]
                    product_dto.product_name = product_name

                    product_dto.to_print()

                    claim_data.append({"claim_number": order_number, "order_number_list": product_dto.get_dict()})

                result = defaultdict(list)

                for item in claim_data:
                    result[item["claim_number"]].append(item["order_number_list"])

                result_dict = {claim_number: order_number_list for claim_number, order_number_list in result.items()}

                order_list = [
                    {"claim_number": claim_number, "order_number_list": order_number_list}
                    for claim_number, order_number_list in result_dict.items()
                ]

            except Exception as e:
                print(str(e))

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_list: {order_list}")

        return order_list

    def order_cancel(self, order):
        driver = self.driver

        claim_number = order["claim_number"]
        order_number_list = order["order_number_list"]

        # 주문번호 이지어드민 검증
        for order_dict in order_number_list:
            try:
                driver.switch_to.window(self.cs_screen_tab)
                check_order_cancel_number_from_ezadmin(self.log_msg, driver, self.shop_name, order_dict)

            except Exception as e:
                print(str(e))
                if self.shop_name in str(e):
                    raise Exception(f"{self.shop_name} {order}: 배송전 주문취소 상태가 아닙니다.")

            finally:
                driver.refresh()
                driver.switch_to.window(self.shop_screen_tab)

        # # 11번가 order_cancel_number -> ordPrdCnSeq, ordNo, ordPrdSeq (클레임번호, 주문번호, 주문순번) 세개의 정보가 담겨있음
        # # 주문번호 이지어드민 검증
        # ordPrdCnSeq = order["ordPrdCnSeq"]
        # ordNo = order["ordNo"]
        # ordPrdSeq = order["ordPrdSeq"]

        try:
            APIBot = ElevenStreetAPI(self.api_key)

            driver.get("https://soffice.11st.co.kr/view/6209?preViewCode=D")

            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//iframe[@title="취소관리"]')))
            time.sleep(0.5)

            for order_dict in order_number_list:
                ordPrdCnSeq = order_dict["상품명"]
                ordNo = order_dict["주문번호"]
                ordPrdSeq = order_dict["주문상세번호"]

                ResultOrder = asyncio.run(APIBot.cancelreqconf_from_ordInfo(ordPrdCnSeq, ordNo, ordPrdSeq))

                if ResultOrder["ResultOrder"]["result_code"] != "0":
                    self.log_msg.emit(f"{self.shop_name} {order_dict}: 취소 승인 메시지를 찾지 못했습니다.")
                    continue

                self.log_msg.emit(f"{self.shop_name} {order_dict}: 취소 완료")

        except Exception as e:
            print(str(e))
            if self.shop_name in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{self.shop_name} {order}: 해당 주문이 존재하지 않습니다.")

    def work_start(self):
        driver = self.driver
        print(f"ElevenStreet: work_start")

        try:
            driver.execute_script(f"window.open('{self.login_url}');")
            self.shop_screen_tab = driver.window_handles[1]
            driver.switch_to.window(self.shop_screen_tab)

            self.login()

            order_list = self.get_order_list()

            self.log_msg.emit(f"{self.shop_name}: {len(order_list)}개의 주문번호(묶음번호)를 발견했습니다.")

            for order in order_list:
                try:
                    self.order_cancel(order)
                except Exception as e:
                    print(str(e))
                    continue

        except Exception as e:
            print(str(e))
            print(f"{self.shop_name} 작업 실패")

        finally:
            driver.close()
            driver.switch_to.window(self.cs_screen_tab)

            self.log_msg.emit(f"{self.shop_name}: 작업 종료")
