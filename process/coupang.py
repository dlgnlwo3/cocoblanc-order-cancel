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


class Coupang:
    def __init__(self, log_msg, driver, cs_screen_tab, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.cs_screen_tab: str = cs_screen_tab
        self.dict_account: dict = dict_account

        self.shop_name = "쿠팡"
        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]

        self.default_wait = 10

    def login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 바로 스토어 화면으로 이동합니다.
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '//h1[contains(text(), "coupang")]'))
            )
            time.sleep(2)

        except Exception as e:
            pass

        try:
            driver.implicitly_wait(1)

            id_input = driver.find_element(By.XPATH, '//input[@id="username"]')
            time.sleep(0.2)
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(self.login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@id="password"]')
            time.sleep(0.2)
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(self.login_pw)

            login_button = driver.find_element(By.XPATH, '//input[contains(@id, "login")]')
            time.sleep(0.2)
            login_button.click()
            time.sleep(0.2)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Coupang Wing")]'))
            )
            time.sleep(0.2)

        except Exception as e:
            print(e)
            self.log_msg.emit(f"{self.shop_name} 로그인 실패")
            raise Exception(f"{self.shop_name} 로그인 실패")

    def get_order_list(self):
        driver = self.driver

        order_list = []
        try:
            driver.get("https://wing.coupang.com/tenants/sfl-portal/stop-shipment/list")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "출고중지관리")]'))
            )
            time.sleep(0.5)

            # 페이지를 넘겨가며 주문번호를 수집해야 함 (한 페이지에 10개씩 출력)
            # 페이지 목록
            # $x('//span[contains(@data-wuic-attrs, "page")]//a')
            last_page_element = driver.find_elements(By.XPATH, '//span[contains(@data-wuic-attrs, "page")]//a')[-1]
            last_page = last_page_element.get_attribute("textContent")
            last_page = last_page.strip()
            last_page = int(last_page)

            print(f"last_page: {last_page}")

            driver.execute_script("arguments[0].click();", last_page_element)
            time.sleep(1)

            claim_data = []
            for num in range(last_page, 0, -1):
                print(f"page_num: {num}")
                try:
                    driver.implicitly_wait(1)
                    num_page_link = driver.find_element(
                        By.XPATH, f'//span[contains(@data-wuic-attrs, "page:{num}")]//a'
                    )
                except Exception as e:
                    try:
                        prev_button = driver.find_element(By.XPATH, '//span[@data-wuic-partial="prev"]//a')
                        driver.execute_script("arguments[0].click();", prev_button)
                        time.sleep(1)
                        num_page_link = driver.find_element(
                            By.XPATH, f'//span[contains(@data-wuic-attrs, "page:{num}")]//a'
                        )
                    except Exception as e:
                        print(str(e))
                finally:
                    driver.implicitly_wait(self.default_wait)

                driver.execute_script("arguments[0].click();", num_page_link)
                time.sleep(1)

                claim_tr_list = driver.find_elements(
                    By.XPATH, '//table[.//tr[./th[contains(text(), "출고중지 처리")]]]//tr[not(th)]'
                )

                for claim_tr in reversed(claim_tr_list):
                    product_dto = ProductDto()

                    # 주문번호
                    order_number = (
                        claim_tr.find_element(
                            By.XPATH,
                            f"./td[12]",
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.order_number = order_number

                    search_product_name_list = claim_tr.find_elements(
                        By.XPATH,
                        f"./td[6]//span",
                    )
                    product_name_list = []
                    for search_product_name in search_product_name_list:
                        product_name = search_product_name.get_attribute("textContent")
                        product_name = product_name[product_name.find("(") + 1 : product_name.rfind(")")]
                        product_name_list.append(product_name)
                    product_dto.product_name = product_name_list

                    product_recv_name = (
                        claim_tr.find_element(
                            By.XPATH,
                            f"./td[9]",
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.product_recv_name = product_recv_name

                    product_recv_tel = (
                        claim_tr.find_element(
                            By.XPATH,
                            f"./td[10]",
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.product_recv_tel = product_recv_tel

                    product_dto.to_print()

                    claim_data.insert(0, {"claim_number": order_number, "order_number_list": product_dto.get_dict()})

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
            driver.get("https://wing.coupang.com/tenants/sfl-portal/stop-shipment/list")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "출고중지관리")]'))
            )
            time.sleep(0.5)

            # 검색 열기
            open_search_button = driver.find_element(By.XPATH, '//button[contains(text(), "검색 열기")]')
            driver.execute_script("arguments[0].click();", open_search_button)
            time.sleep(0.2)

            # 검색조건 -> 주문번호
            search_option_select = Select(driver.find_element(By.XPATH, '//select[./option[contains(text(), "주문번호")]]'))
            search_option_select.select_by_value("orderId")
            time.sleep(0.2)

            # 주문번호 입력
            input_keyword = driver.find_element(By.XPATH, '//span[contains(@class, "search-words-input")]//input')
            input_keyword.send_keys(claim_number)
            time.sleep(0.2)

            # 검색 클릭
            search_submit_button = driver.find_element(By.XPATH, '//button[contains(text(), "검색") and @type="submit"]')
            driver.execute_script("arguments[0].click();", search_submit_button)
            time.sleep(2)

            # 취소 품목 체크박스
            # '//tr[./td[.//span[contains(text(), "2023061616868779870")]]]//input[@type="checkbox"]'
            try:
                order_cancel_target_checkbox = driver.find_element(
                    By.XPATH, f'//tr[.//a[contains(text(), "{claim_number}")]]/td//input[@type="checkbox"]'
                )
                driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
                time.sleep(0.2)
            except Exception as e:
                print(str(e))
                self.log_msg.emit(f"{self.shop_name} {order}: 검색 결과가 없습니다.")
                raise Exception(f"{self.shop_name} {order}: 검색 결과가 없습니다.")

            # 출고중지완료 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[contains(text(), "출고중지완료")]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # modal창: 하기 1건을 출고중지 완료 하시겠습니까? 출고중지완료하시면 환불이 완료됩니다. or 상품을 먼저 선택해 주세요.
            # $x('//div[@data-wuic-partial="widget"][.//span[contains(text(), "출고중지완료하시면 환불이 완료됩니다.")]]')
            try:
                coupang_order_cancel_button = driver.find_element(
                    By.XPATH,
                    '//div[@data-wuic-partial="widget"][.//span[contains(text(), "출고중지완료하시면 환불이 완료됩니다.")]]//div[@class="footer"]/button[contains(text(), "완료")]',
                )
                driver.execute_script("arguments[0].click();", coupang_order_cancel_button)
                time.sleep(2)

                # 별다른 메시지 없이 처리됨
                self.log_msg.emit(f"{self.shop_name} {order}: 취소 완료")

            except Exception as e:
                print(str(e))
                self.log_msg.emit(f"{self.shop_name} {order}: 취소 성공 메시지를 찾지 못했습니다.")
                raise Exception(f"{self.shop_name} {order}: 취소 성공 메시지를 찾지 못했습니다.")

        except Exception as e:
            print(str(e))
            if self.shop_name in str(e):
                raise Exception(str(e))

    def check_order_cancel_number_from_ezadmin(self, order_dict: dict):
        driver = self.driver

        order_number = order_dict["주문번호"]
        product_name_list = order_dict["상품명"]
        product_recv_name = order_dict["수령자명"]
        product_recv_tel = order_dict["수령자연락처"]

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
                        if not (search_product_option in product_name_list):
                            continue

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
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        search_product_recv_tel = (
                            driver.find_element(By.XPATH, '//td[@id="di_recv_tel"]')
                            .get_attribute("textContent")
                            .strip()
                        )
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
        print(f"Coupang: work_start")

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
