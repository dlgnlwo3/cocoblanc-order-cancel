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


class KakaoTalkStore:
    def __init__(self, log_msg, driver, cs_screen_tab, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.cs_screen_tab: str = cs_screen_tab
        self.dict_account: dict = dict_account

        self.shop_name = "카카오톡스토어"
        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]

        self.default_wait = 10

    def login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 해당 web element가 존재하지 않습니다.
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '//button[@type="submit"][contains(text(), "로그인")]'))
            )
            time.sleep(1)

        except Exception as e:
            print(f"{self.shop_name} 로그인 화면이 아닙니다.")

        try:
            driver.implicitly_wait(1)

            login_id = self.dict_account["ID"]
            login_pw = self.dict_account["PW"]

            id_input = driver.find_element(By.XPATH, '//input[@name="loginKey"]')
            time.sleep(0.2)
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@name="password"]')
            time.sleep(0.2)
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(login_pw)

            login_button = driver.find_element(By.XPATH, '//button[@type="submit"][contains(text(), "로그인")]')
            login_button.click()
            time.sleep(1)

        except Exception as e:
            print(f"{self.shop_name} 로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//h1[./a[./img[@alt="톡스토어 판매자센터"]]]'))
            )
            time.sleep(0.5)

        except Exception as e:
            self.log_msg.emit(f"{self.shop_name} 로그인 실패")
            raise Exception(f"{self.shop_name} 로그인 실패")

        # 각종 팝업창 닫기
        try:
            popup_close_button = driver.find_element(
                By.XPATH, '//div[@class="popup-foot"]//button[./span[contains(text(), "닫기")]]'
            )
            driver.execute_script("arguments[0].click();", popup_close_button)
            time.sleep(0.2)

        except Exception as e:
            print("popup not found")

    def get_order_list(self):
        driver = self.driver

        order_list = []
        try:
            driver.get("https://store-buy-sell.kakao.com/order/cancelList?orderSummaryCount=CancelRequestToBuyer")

            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "구매자 취소 요청")]'))
            )
            time.sleep(0.2)

            # 500개씩 보기
            pageSize_select = Select(driver.find_element(By.XPATH, '//select[@name="pageSize"]'))
            pageSize_select.select_by_visible_text("500개씩")
            time.sleep(1)

            try:
                driver.implicitly_wait(1)
                not_found_message = driver.find_element(By.XPATH, '//div[text()="목록이 없습니다."]').get_attribute(
                    "textContent"
                )
                print(not_found_message)
                return order_list

            except Exception as e:
                pass

            finally:
                driver.implicitly_wait(self.default_wait)

            # 클레임번호 목록
            first_line_in_table = driver.find_element(
                By.XPATH,
                '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding"))]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_1_")]',
            )
            driver.execute_script("arguments[0].click();", first_line_in_table)
            time.sleep(0.2)

            # 목록의 갯수
            claim_count = driver.find_element(By.XPATH, '//div[contains(@id, "AX_gridStatus")]/b').get_attribute(
                "textContent"
            )
            if claim_count.isdigit():
                claim_count = int(claim_count)
            else:
                claim_count = 0

            claim_data = []
            for i in range(1, claim_count + 1):
                # 현재 활성화 된 tr
                # $x('//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]')
                product_dto = ProductDto()

                claim_number = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_1_")]',
                ).get_attribute("textContent")

                order_number = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_3_")]',
                ).get_attribute("textContent")
                product_dto.order_number = order_number

                product_name = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_14_")]',
                ).get_attribute("textContent")
                product_dto.product_name = product_name

                product_option = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_15_")]',
                ).get_attribute("textContent")
                product_option = product_option.replace(": ", "/").replace(", ", ",")
                product_dto.product_option = product_option

                product_qty = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_16_")]',
                ).get_attribute("textContent")
                product_dto.product_qty = product_qty

                product_recv_name = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_25_")]',
                ).get_attribute("textContent")
                product_dto.product_recv_name = product_recv_name

                product_recv_tel = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_26_")]',
                ).get_attribute("textContent")
                product_dto.product_recv_tel = product_recv_tel

                product_dto.to_print()

                claim_data.append({"claim_number": claim_number, "order_number_list": product_dto.get_dict()})

                send_keys_to_driver(driver, Keys.ARROW_DOWN)

            time.sleep(0.2)

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
                    raise Exception(f"{str(e)}")

            finally:
                driver.refresh()
                driver.switch_to.window(self.shop_screen_tab)

        try:
            driver.get("https://store-buy-sell.kakao.com/order/cancelList?orderSummaryCount=CancelRequestToBuyer")

            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "구매자 취소 요청")]'))
            )
            time.sleep(0.2)

            # 주문번호 입력
            input_orderIdList = driver.find_element(By.XPATH, '//input[@name="orderIdList"]')
            input_orderIdList.send_keys(order_dict["주문번호"])
            time.sleep(0.2)

            # 검색 클릭
            search_button = driver.find_element(By.XPATH, '//button[@type="submit" and text()="검색"]')
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(1)

            # 취소 품목
            order_cancel_target = driver.find_element(
                By.XPATH,
                f'//button[contains(@onclick, "claim.popOrderDetail") and contains(@onclick, "{order_dict["주문번호"]}")]',
            )
            driver.execute_script("arguments[0].click();", order_cancel_target)
            time.sleep(1)

            # 새 창 열림
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(order_cancel_tab)
                time.sleep(1)

                order_cancel_iframe = driver.find_element(By.XPATH, '//iframe[contains(@src, "omsOrderDetail")]')
                driver.switch_to.frame(order_cancel_iframe)
                time.sleep(0.5)

                order_cancel_button = driver.find_element(By.XPATH, '//button[contains(text(), "취소승인(환불)")]')
                driver.execute_script("arguments[0].click();", order_cancel_button)
                time.sleep(0.5)

                # 취소 승인 하시겠습니까? alert
                alert_msg = ""
                try:
                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    alert = driver.switch_to.alert
                    alert_msg = alert.text
                except Exception as e:
                    print(f"no alert")

                print(f"{alert_msg}")

                if "취소 승인 하시겠습니까" in alert_msg:
                    alert.accept()

                    # 정상처리 되었습니다. alert
                    try:
                        WebDriverWait(driver, 10).until(EC.alert_is_present())
                    except Exception as e:
                        print(f"no alert")
                        self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                        raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                    alert_ok_try(driver)

                elif alert_msg != "":
                    alert.accept()
                    self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")
                    raise Exception(f"{self.shop_name} {order}: {alert_msg}")

                else:
                    self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                    raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                self.log_msg.emit(f"{self.shop_name} {order}: 취소 완료")

            except Exception as e:
                print(str(e))
                if self.shop_name in str(e):
                    raise Exception(str(e))

            finally:
                driver.close()
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))
            if self.shop_name in str(e):
                raise Exception(str(e))

    def work_start(self):
        driver = self.driver
        print(f"KakaoTalkStore: work_start")

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
