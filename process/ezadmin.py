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


class Ezadmin:
    def __init__(self, log_msg, driver, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.dict_account: dict = dict_account

        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]

        self.default_wait = 10

    def login(self):
        driver = self.driver
        self.driver.get(self.login_url)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//body[@class="ezadmin-main-body"]'))
        )
        time.sleep(0.2)

        # 로그인 시도
        # 이 행위 중 하나라도 실패한다면 로그인 실패
        try:
            driver.implicitly_wait(2)

            open_button = driver.find_element(By.XPATH, '//a[./span[@class="img_login"]][contains(text(), "로그인")]')
            driver.execute_script("arguments[0].click();", open_button)
            time.sleep(0.2)

            domain_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-domain"]')
            domain_input.clear()
            time.sleep(0.2)
            domain_input.send_keys(self.login_domain)
            time.sleep(0.2)

            id_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-id"]')
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(self.login_id)
            time.sleep(0.2)

            pwd_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-pwd"]')
            pwd_input.clear()
            time.sleep(0.2)
            pwd_input.send_keys(self.login_pw)
            time.sleep(0.2)

            save_domain = driver.find_element(By.XPATH, '//input[@id="savedomain"]')
            driver.execute_script("arguments[0].click();", save_domain)
            time.sleep(0.2)

            login_button = driver.find_element(By.XPATH, '//input[@class="login-btn" and @value="로그인"]')
            driver.execute_script("arguments[0].click();", login_button)
            time.sleep(0.2)

            # 로그인 성공 시 나오는 화면
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[contains(text(), "코코블랑")]'))
                )
            except Exception as e:
                print(e)

            time.sleep(0.2)

            self.close_ezadmin_notice_popups()

        except Exception as e:
            self.log_msg.emit(f"이지어드민 로그인 실패")
            print(e)
            raise Exception(f"이지어드민 로그인 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

    def close_ezadmin_notice_popups(self):
        driver = self.driver
        try:
            driver.implicitly_wait(1)

            driver.execute_script("hide_board('internal_board');")
            time.sleep(0.2)

            driver.execute_script("hide_board('sys_notice_board');")
            time.sleep(0.2)

            # $x('//a[contains(text(), "팝업 전체 닫기")]')
            close_all_popups = driver.find_element(By.XPATH, '//a[contains(text(), "팝업 전체 닫기")]')
            driver.execute_script("arguments[0].click();", close_all_popups)
            time.sleep(0.2)

            # 부가서비스 기간 만료 안내
            try:
                close_service_notice_button = driver.find_element(
                    By.XPATH,
                    '//div[@class="modal-dialog"][.//*[contains(text(), "부가서비스 기간 만료 안내")]]//a[@class="page-close"]',
                )
                driver.execute_script("arguments[0].click();", close_service_notice_button)
                time.sleep(0.2)

            except Exception as e:
                print("부가서비스 기간 만료 안내창 없음")

        except Exception as e:
            print(e)

        finally:
            driver.implicitly_wait(self.default_wait)

    def switch_to_cs_screen(self):
        driver = self.driver

        # cs창을 여는 javascript, 새 창에서 열리게 된다.
        driver.execute_script("javascript:popup_new_cs();")
        time.sleep(1)

        # driver.window_handles
        tabs = driver.window_handles

        # 새 창 (cs창)이 열리지 않았다면 작업 실패
        if len(tabs) == 1:
            raise Exception("cs창에 진입하지 못했습니다.")

        # 새 창이 열렸다면 기존 창을 닫고 tab 전환
        try:
            driver.close()

        except Exception as e:
            print(str(e))

        finally:
            driver.switch_to.window(driver.window_handles[-1])

        # 현재 활성화 된 창이 cs창인지 검증
        try:
            cs_screen_title = driver.find_element(By.XPATH, '//span[@class="pgtitle"][text()="C/S"]')

        except Exception as e:
            print(str(e))
            raise Exception("cs창이 활성화되지 않았습니다.")

        # cs_screen_tab
        return driver.window_handles[0]


def check_order_cancel_number_from_ezadmin(
    log_msg: SignalInstance, driver: webdriver.Chrome, shop_name: str, order_dict: dict
):
    is_checked = False

    order_number = order_dict["주문번호"]
    order_detail_number = order_dict["주문상세번호"]
    product_name = order_dict["상품명"]
    product_option: str = order_dict["상품옵션"]
    product_qty = order_dict["수량"]
    product_recv_name = order_dict["수령자명"]
    product_recv_tel = order_dict["수령자연락처"]

    try:
        if shop_name == "네이버" or shop_name == "지그재그":
            input_tel_number = driver.find_element(By.XPATH, '//td[contains(text(), "전화번호")]/input')
            input_tel_number.clear()
            input_tel_number.send_keys(product_recv_tel)
        else:
            input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "주문번호")]/input')
            input_order_number.clear()
            input_order_number.send_keys(order_number)

        search_button = driver.find_element(By.XPATH, '//div[@id="search"][text()="검색"]')
        driver.execute_script("arguments[0].click();", search_button)
        time.sleep(2)

        grid_order_trs = driver.find_elements(By.XPATH, '//table[@id="grid_order"]//tr[not(@class="jqgfirstrow")]')

        if len(grid_order_trs) == 0:
            log_msg.emit(f"{shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")
            raise Exception(f"{shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")

        for grid_order_tr in grid_order_trs:
            try:
                driver.execute_script("arguments[0].click();", grid_order_tr)
                time.sleep(0.2)

                grid_product_trs = driver.find_elements(
                    By.XPATH,
                    f'//table[contains(@id, "grid_product")]//td[contains(@title, "list_order_id") and contains(@title, "{order_number}")]',
                )

                if len(grid_product_trs) == 0:
                    log_msg.emit(f"{shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")
                    raise Exception(f"{shop_name} {order_dict}: 이지어드민 검색 결과가 없습니다.")

                for grid_product_tr in grid_product_trs:
                    driver.execute_script("arguments[0].click();", grid_product_tr)
                    time.sleep(0.2)

                    search_product_name = (
                        driver.find_element(By.XPATH, '//td[@id="di_shop_pname"]').get_attribute("textContent").strip()
                    )

                    search_product_option = (
                        driver.find_element(By.XPATH, '//td[@id="di_shop_options"]')
                        .get_attribute("textContent")
                        .strip()
                    )

                    search_product_qty = (
                        driver.find_element(By.XPATH, '//td[@id="di_order_qty"]').get_attribute("textContent").strip()
                    )

                    search_order_number = (
                        driver.find_element(By.XPATH, '//td[@id="di_order_id"]').get_attribute("textContent").strip()
                    )

                    search_order_detail_number = (
                        driver.find_element(By.XPATH, '//td[@id="di_order_id_seq"]')
                        .get_attribute("textContent")
                        .strip()
                    )

                    search_product_recv_name = (
                        driver.find_element(By.XPATH, '//td[@id="di_recv_name"]').get_attribute("textContent").strip()
                    )

                    search_product_recv_tel = (
                        driver.find_element(By.XPATH, '//td[@id="di_recv_tel"]').get_attribute("textContent").strip()
                    )

                    search_product_cs_state = (
                        driver.find_element(By.XPATH, '//td[@id="di_product_cs"]').get_attribute("textContent").strip()
                    )

                    # 각 판매처마다 검증해야하는 항목이 다릅니다.
                    if shop_name == "카카오톡스토어":
                        # 판매처 상품명
                        if not (product_name in search_product_name):
                            continue

                        # 판매처 옵션
                        if not (product_option in search_product_option):
                            continue

                        # 주문수량
                        if not (product_qty in search_product_qty):
                            continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # # 주문상세번호
                        # if not (order_detail_number in search_order_detail_number):
                        #     continue

                        # # 수령자
                        # if not (product_recv_name in search_product_recv_name):
                        #     continue

                        # # 수령자 연락처
                        # if not (product_recv_tel in search_product_recv_tel):
                        #     continue

                    elif shop_name == "위메프":
                        # 판매처 상품명
                        if not (product_name in search_product_name):
                            continue

                        # 판매처 옵션
                        if not (product_option in search_product_option):
                            continue

                        # 주문수량
                        if not (product_qty in search_product_qty):
                            continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # # 주문상세번호
                        # if not (order_detail_number in search_order_detail_number):
                        #     continue

                        # 수령자
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        # # 수령자 연락처
                        # if not (product_recv_tel in search_product_recv_tel):
                        #     continue

                    elif shop_name == "티몬":
                        # 판매처 상품명
                        if not (product_name in search_product_name):
                            continue

                        # 판매처 옵션
                        if not (product_option in search_product_option):
                            continue

                        # 주문수량
                        if not (product_qty in search_product_qty):
                            continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # 주문상세번호
                        if not (order_detail_number in search_order_detail_number):
                            continue

                        # 수령자
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        # 수령자 연락처
                        if not (product_recv_tel in search_product_recv_tel):
                            continue

                    elif shop_name == "지그재그":
                        # 판매처 상품명
                        if not (product_name in search_product_name):
                            continue

                        # 판매처 옵션
                        if not (product_option in search_product_option):
                            continue

                        # 주문수량
                        if not (product_qty in search_product_qty):
                            continue

                        # # 주문번호
                        # if not (order_number in search_order_number):
                        #     continue

                        # 주문상세번호
                        if not (order_number in search_order_detail_number):
                            continue

                        # 수령자
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        # 수령자 연락처
                        if not (product_recv_tel in search_product_recv_tel):
                            continue

                    elif shop_name == "브리치":
                        # # 판매처 상품명
                        # if not (product_name in search_product_name):
                        #     continue

                        # # 판매처 옵션
                        # if not (product_option in search_product_option):
                        #     continue

                        # # 주문수량
                        # if not (product_qty in search_product_qty):
                        #     continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # 주문상세번호
                        if not (order_detail_number in search_order_detail_number):
                            continue

                        # # 수령자
                        # if not (product_recv_name in search_product_recv_name):
                        #     continue

                        # # 수령자 연락처
                        # if not (product_recv_tel in search_product_recv_tel):
                        #     continue

                    elif shop_name == "쿠팡":
                        # # 판매처 상품명
                        # if not (search_product_name in product_name):
                        #     continue

                        # 판매처 옵션
                        if not (search_product_option in product_name):
                            continue

                        # # 주문수량
                        # if not (product_qty in search_product_qty):
                        #     continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # # 주문상세번호
                        # if not (order_detail_number in search_order_detail_number):
                        #     continue

                        # 수령자
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        # # 수령자 연락처
                        # if not (product_recv_tel in search_product_recv_tel):
                        #     continue

                    elif shop_name == "11번가":
                        # # 판매처 상품명
                        # if not (product_name in search_product_name):
                        #     continue

                        # 판매처 옵션
                        if not (search_product_option in product_option):
                            continue

                        # 주문수량
                        if not (product_qty in search_product_qty):
                            continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # 주문상세번호
                        if not (order_detail_number in search_order_detail_number):
                            continue

                        # # 수령자
                        # if not (product_recv_name in search_product_recv_name):
                        #     continue

                        # # 수령자 연락처
                        # if not (product_recv_tel in search_product_recv_tel):
                        #     continue

                    elif shop_name == "네이버":
                        # 판매처 상품명
                        if not (product_name in search_product_name):
                            continue

                        # 판매처 옵션
                        if not (product_option in search_product_option):
                            continue

                        # 주문수량
                        if not (product_qty in search_product_qty):
                            continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # 주문상세번호
                        if not (order_detail_number in search_order_detail_number):
                            continue

                        # 수령자
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        # 수령자 연락처
                        if not (product_recv_tel in search_product_recv_tel):
                            continue

                    else:
                        # 판매처 상품명
                        if not (product_name in search_product_name):
                            continue

                        # 판매처 옵션
                        if not (product_option in search_product_option):
                            continue

                        # 주문수량
                        if not (product_qty in search_product_qty):
                            continue

                        # 주문번호
                        if not (order_number in search_order_number):
                            continue

                        # 주문상세번호
                        if not (order_detail_number in search_order_detail_number):
                            continue

                        # 수령자
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        # 수령자 연락처
                        if not (product_recv_tel in search_product_recv_tel):
                            continue

                    # C/S 상태
                    if not "배송전 주문취소" in search_product_cs_state:
                        is_checked = False
                        log_msg.emit(f"{shop_name} {order_dict}: 배송전 주문취소 상태가 아닙니다.")
                        raise Exception(f"{shop_name} {order_dict}: 배송전 주문취소 상태가 아닙니다.")

                    log_msg.emit(
                        f"{search_order_number}: {search_product_name}, {search_product_option}, {search_product_qty}, {search_product_recv_name}, {search_product_recv_tel}, {search_product_cs_state}"
                    )
                    print(
                        f"{search_order_number}: {search_product_name}, {search_product_option}, {search_product_qty}, {search_product_recv_name}, {search_product_recv_tel}, {search_product_cs_state}"
                    )

                    is_checked = True

                    if is_checked:
                        return
                    else:
                        print(f"알 수 없는 오류")
                        raise Exception(f"{shop_name} {order_dict}: 알 수 없는 오류가 발생했습니다.")

            except Exception as e:
                print(str(e))
                if shop_name in str(e):
                    raise Exception(str(e))

            finally:
                tab_close_button = driver.find_element(By.XPATH, '//span[contains(@class, "ui-icon-close")]')
                driver.execute_script("arguments[0].click();", tab_close_button)
                time.sleep(0.2)

        if not is_checked:
            log_msg.emit(f"{shop_name} {order_dict}: 일치하는 정보가 없습니다.")
            raise Exception(f"{shop_name} {order_dict}: 일치하는 정보가 없습니다.")

    except Exception as e:
        print(str(e))
        if shop_name in str(e):
            raise Exception(str(e))
