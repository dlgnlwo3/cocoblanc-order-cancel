if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from selenium import webdriver
from dtos.product_dto import ProductDto

from PySide6.QtCore import SignalInstance

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


class Bflow:
    def __init__(self, log_msg, driver, cs_screen_tab, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.cs_screen_tab: str = cs_screen_tab
        self.dict_account: dict = dict_account

        self.shop_name = "브리치"
        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]

        self.default_wait = 10

    def login(self):
        driver = self.driver

        try:
            WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH, '//button[text()="로그인"]')))
            time.sleep(1)

            login_button = driver.find_element(By.XPATH, '//button[text()="로그인"]')
            driver.execute_script("arguments[0].click();", login_button)
            time.sleep(0.2)

        except Exception as e:
            print("로그인 화면이 아닙니다.")

        try:
            driver.implicitly_wait(1)

            id_input = driver.find_element(
                By.XPATH, '//div[@class="login-area"]//input[@class="login-input"][@placeholder="이메일"]'
            )
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(self.login_id)

            pw_input = driver.find_element(
                By.XPATH, '//div[@class="login-area"]//input[@class="login-input"][@placeholder="비밀번호"]'
            )
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(self.login_pw)

            login_button = driver.find_element(By.XPATH, '//button[./span[contains(text(), "로그인")]]')
            time.sleep(0.2)
            login_button.click()
            time.sleep(0.2)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            driver.implicitly_wait(5)
            main_page = driver.find_element(By.XPATH, '//div[@id="main-page"]')

            time.sleep(5)

        except Exception as e:
            self.log_msg.emit(f"{self.shop_name} 로그인 실패")
            print(str(e))
            raise Exception(f"{self.shop_name} 로그인 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        # 팝업 닫기
        # $x('//div[contains(@class, "modal")]//span[@class="close-btn"]')
        try:
            driver.implicitly_wait(1)

            modal_close_btn = driver.find_element(
                By.XPATH, '//div[contains(@class, "modal")]//span[@class="close-btn"]'
            )
            driver.execute_script("arguments[0].click();", modal_close_btn)
            time.sleep(0.2)

        except Exception as e:
            print("no modal popup")

        finally:
            driver.implicitly_wait(self.default_wait)

    def get_order_list(self):
        driver = self.driver

        order_list = []
        try:
            driver.get("https://b-flow.co.kr/order/cancels")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "취소 관리")]'))
            )
            time.sleep(0.2)

            wait_loading(driver, '//div[contains(@class, "overlay")]')

            # 취소요청 x 건 클릭
            # $x('//span[contains(text(), "취소요청")]/span[contains(@class, "text-link") and contains(text(), "건")]')
            cancel_request_link = driver.find_element(
                By.XPATH,
                '//span[contains(text(), "취소요청")]/span[contains(@class, "text-link") and contains(text(), "건")]',
            )
            driver.execute_script("arguments[0].click();", cancel_request_link)
            wait_loading(driver, '//div[contains(@class, "overlay")]')
            time.sleep(0.2)

            # 500개씩 보기
            # $x('//span[contains(@class, "option")][./span[text()="500"]]')
            option_span = driver.find_element(By.XPATH, '//span[contains(@class, "option")][./span[text()="500"]]')
            driver.execute_script("arguments[0].click();", option_span)
            wait_loading(driver, '//div[contains(@class, "overlay")]')
            time.sleep(0.2)

            # 데이터 목록
            # $x('//table[contains(@class, "data-table")]//tbody/tr')
            claim_tr_list = driver.find_elements(By.XPATH, '//table[contains(@class, "data-table")]//tbody/tr')

            claim_data = []
            for claim_tr in claim_tr_list:
                product_dto = ProductDto()

                # 주문번호 (묶음번호)
                claim_number = (
                    claim_tr.find_element(
                        By.XPATH,
                        f"./td[2]//span",
                    )
                    .get_attribute("textContent")
                    .strip()
                )
                product_dto.order_number = claim_number

                # 주문상세번호 (개별번호)
                order_number = (
                    claim_tr.find_element(
                        By.XPATH,
                        f"./td[3]/div",
                    )
                    .get_attribute("textContent")
                    .strip()
                )
                product_dto.order_detail_number = order_number

                product_dto.to_print()

                claim_data.append({"claim_number": claim_number, "order_number_list": product_dto.get_dict()})

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
                self.check_order_cancel_number_from_ezadmin(order_dict)

            except Exception as e:
                print(str(e))
                if self.shop_name in str(e):
                    raise Exception(f"{self.shop_name} {order}: 배송전 주문취소 상태가 아닙니다.")

        try:
            driver.get("https://b-flow.co.kr/order/cancels")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "취소 관리")]'))
            )
            time.sleep(0.2)

            wait_loading(driver, '//div[contains(@class, "overlay")]')

            # 취소요청 x 건 클릭
            # $x('//span[contains(text(), "취소요청")]/span[contains(@class, "text-link") and contains(text(), "건")]')
            cancel_request_link = driver.find_element(
                By.XPATH,
                '//span[contains(text(), "취소요청")]/span[contains(@class, "text-link") and contains(text(), "건")]',
            )
            driver.execute_script("arguments[0].click();", cancel_request_link)
            wait_loading(driver, '//div[contains(@class, "overlay")]')
            time.sleep(0.2)

            # 500개씩 보기
            # $x('//span[contains(@class, "option")][./span[text()="500"]]')
            option_span = driver.find_element(By.XPATH, '//span[contains(@class, "option")][./span[text()="500"]]')
            driver.execute_script("arguments[0].click();", option_span)
            wait_loading(driver, '//div[contains(@class, "overlay")]')
            time.sleep(0.2)

            # 취소 품목 체크박스
            # '//tr[./td[.//span[contains(text(), "2023061616868779870")]]]//input[@type="checkbox"]'
            order_cancel_target_checkbox = driver.find_element(
                By.XPATH, f'//tr[./td[.//span[contains(text(), "{claim_number}")]]]//input[@type="checkbox"]'
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(1)

            # 취소완료 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[contains(text(), "취소완료")]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # 1개의 항목을 취소완료 처리하시겠습니까? alert
            alert_msg = ""
            try:
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert_msg = alert.text
            except Exception as e:
                print(f"no alert")
                self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

            print(f"{alert_msg}")

            if "취소완료 처리하시겠습니까" in alert_msg:
                alert.accept()
                time.sleep(1)

                # [취소번호: xxxxxxxx] 완료 처리되었습니다.
                try:
                    cancel_success_message = driver.find_element(
                        By.XPATH, '//li[contains(text(), "완료 처리되었습니다") and contains(text(), "취소번호")]'
                    ).get_attribute("textContent")
                    print(cancel_success_message)

                except Exception as e:
                    print(str(e))
                    self.log_msg.emit(f"{self.shop_name} {order}: 취소 성공 메시지를 찾지 못했습니다.")
                    raise Exception(f"{self.shop_name} {order}: 취소 성공 메시지를 찾지 못했습니다.")

            elif alert_msg != "":
                alert.dismiss()
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

    def check_order_cancel_number_from_ezadmin(self, order_dict: dict):
        driver = self.driver

        order_number = order_dict["주문번호"]
        order_detail_number = order_dict["주문상세번호"]

        try:
            driver.switch_to.window(self.cs_screen_tab)

            input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "주문번호")]/input')

            input_order_number.clear()

            input_order_number.send_keys(order_number)

            search_button = driver.find_element(By.XPATH, '//div[@id="search"][text()="검색"]')
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(2)

            grid_order_trs = driver.find_elements(By.XPATH, '//table[@id="grid_order"]//tr[not(@class="jqgfirstrow")]')

            if len(grid_order_trs) == 0:
                self.log_msg.emit(f"{self.shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")
                raise Exception(f"{self.shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")

            for grid_order_tr in grid_order_trs:
                try:
                    driver.execute_script("arguments[0].click();", grid_order_tr)
                    time.sleep(0.2)

                    grid_product_trs = driver.find_elements(
                        By.XPATH,
                        f'//table[contains(@id, "grid_product")]//td[contains(@title, "list_order_id") and contains(@title, "{order_number}")]',
                    )

                    if len(grid_product_trs) == 0:
                        self.log_msg.emit(f"{self.shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")
                        raise Exception(f"{self.shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")

                    for grid_product_tr in grid_product_trs:
                        driver.execute_script("arguments[0].click();", grid_product_tr)
                        time.sleep(0.2)

                        search_product_name = (
                            driver.find_element(By.XPATH, '//td[@id="di_shop_pname"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        search_product_option = (
                            driver.find_element(By.XPATH, '//td[@id="di_shop_options"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        search_order_number = (
                            driver.find_element(By.XPATH, '//td[@id="di_order_id"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if not (order_number in search_order_number):
                            continue

                        search_order_detail_number = (
                            driver.find_element(By.XPATH, '//td[@id="di_order_id_seq"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if not (order_detail_number in search_order_detail_number):
                            continue

                        search_product_qty = (
                            driver.find_element(By.XPATH, '//td[@id="di_order_qty"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        search_product_recv_name = (
                            driver.find_element(By.XPATH, '//td[@id="di_recv_name"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        # search_product_recv_tel = (
                        #     driver.find_element(By.XPATH, '//td[@id="di_recv_tel"]')
                        #     .get_attribute("textContent")
                        #     .strip()
                        # )
                        # if not (product_recv_tel in search_product_recv_tel):
                        #     continue

                        product_cs_state = (
                            driver.find_element(By.XPATH, '//td[@id="di_product_cs"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        if not "배송전 주문취소" in product_cs_state:
                            self.log_msg.emit(f"{self.shop_name} {order_dict}: 배송전 주문취소 상태가 아닙니다.")
                            raise Exception(f"{self.shop_name} {order_dict}: 배송전 주문취소 상태가 아닙니다.")

                        print(
                            f"{search_order_number}: {search_product_name}, {search_product_option}, {search_product_qty}, {search_product_recv_name}, {product_cs_state}"
                        )

                except Exception as e:
                    print(str(e))
                    if self.shop_name in str(e):
                        raise Exception(str(e))

                finally:
                    tab_close_button = driver.find_element(By.XPATH, '//span[contains(@class, "ui-icon-close")]')
                    driver.execute_script("arguments[0].click();", tab_close_button)
                    time.sleep(0.2)

        except Exception as e:
            print(str(e))
            if self.shop_name in str(e):
                raise Exception(str(e))

        finally:
            driver.refresh()
            driver.switch_to.window(self.shop_screen_tab)

    def work_start(self):
        driver = self.driver
        print(f"Bflow: work_start")

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
