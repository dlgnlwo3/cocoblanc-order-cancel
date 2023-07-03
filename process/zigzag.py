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


class Zigzag:
    def __init__(self, log_msg, driver, cs_screen_tab, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.cs_screen_tab: str = cs_screen_tab
        self.dict_account: dict = dict_account

        self.shop_name = "지그재그"
        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]

        self.default_wait = 10

    def login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 바로 스토어 선택 화면으로 이동합니다.
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '//h1[contains(text(), "파트너센터 로그인")]'))
            )
            time.sleep(2)

        except Exception as e:
            pass

        try:
            driver.implicitly_wait(1)

            id_input = driver.find_element(By.XPATH, '//input[contains(@placeholder, "이메일")]')
            id_input.click()
            time.sleep(0.2)
            id_input.send_keys(Keys.LEFT_CONTROL, "a", Keys.BACK_SPACE)
            time.sleep(0.2)
            id_input.send_keys(self.login_id)

            pw_input = driver.find_element(By.XPATH, '//input[contains(@placeholder, "비밀번호")]')
            pw_input.click()
            time.sleep(0.2)
            pw_input.send_keys(Keys.LEFT_CONTROL, "a", Keys.BACK_SPACE)
            time.sleep(0.2)
            pw_input.send_keys(self.login_pw)

            login_button = driver.find_element(By.XPATH, '//button[contains(text(), "로그인")]')
            time.sleep(0.2)
            login_button.click()
            time.sleep(0.2)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[contains(text(), "코코블랑")]'))
            )
            time.sleep(0.2)

            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//span[contains(text(), "광고 관리")]'))
            )
            time.sleep(0.2)

        except Exception as e:
            print(str(e))
            self.log_msg.emit("지그재그 로그인 실패")
            raise Exception("지그재그 로그인 실패")

    def get_order_list(self):
        driver = self.driver

        order_list = []
        try:
            driver.get("https://partners.kakaostyle.com/shop/cocoblanc/order_item/list/cancel")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[text()="취소관리"]')))
            time.sleep(0.2)

            # xx개씩 보기 펼치기
            StyledSelectTextWrapper = driver.find_element(
                By.XPATH, '//div[contains(@class, "StyledSelectTextWrapper")][./span[contains(text(), "씩 보기")]]'
            )
            driver.execute_script("arguments[0].click();", StyledSelectTextWrapper)
            time.sleep(0.2)

            # 500개씩 보기
            StyledMenuItem = driver.find_element(
                By.XPATH, '//div[contains(@class, "StyledMenuItem") and text()="500개씩 보기"]'
            )
            driver.execute_script("arguments[0].click();", StyledMenuItem)
            time.sleep(1)

            # 주문번호 -> 클레임번호 (묶음번호)
            # 묶음번호 목록
            # $x('//tr[contains(@class, "TableRow")]/td[@rowspan]/button')
            claim_number_tr_list = driver.find_elements(
                By.XPATH, '//tr[contains(@class, "TableRow")]/td[@rowspan]/button'
            )

            claim_number_list = []
            for claim_number_tr in claim_number_tr_list:
                claim_number = claim_number_tr.get_attribute("textContent")
                claim_number_list.append(claim_number)

            # 해당 주문번호에 묶여있는 행을 특정할 수 없어서 주문번호로 검색 후 나오는 모든 행을 추출해야 함
            claim_data = []
            for claim_number in claim_number_list:
                # 상세조건 클릭 -> 셀렉트박스
                # $x('//label[./span[text()="상세조건"]]/following-sibling::div//div[contains(@class, "StyledSelectWrapper")]')
                StyledSelectWrapper = driver.find_element(
                    By.XPATH,
                    '//label[./span[text()="상세조건"]]/following-sibling::div//div[contains(@class, "StyledSelectWrapper")]',
                )
                driver.execute_script("arguments[0].click();", StyledSelectWrapper)
                time.sleep(0.2)

                # 주문 번호 클릭 -> 옵션
                # $x('//div[contains(@class, "StyledMenuItem") and contains(text(), "주문 번호")]')
                StyledMenuItem = driver.find_element(
                    By.XPATH, '//div[contains(@class, "StyledMenuItem") and contains(text(), "주문 번호")]'
                )
                driver.execute_script("arguments[0].click();", StyledMenuItem)
                time.sleep(0.2)

                # 주문번호 입력 -> 인풋
                # $x('//label[./span[text()="상세조건"]]/following-sibling::div//input')
                number_input = driver.find_element(
                    By.XPATH, '//label[./span[text()="상세조건"]]/following-sibling::div//input'
                )
                number_input.clear()
                number_input.send_keys(claim_number)
                time.sleep(0.2)

                search_button = driver.find_element(By.XPATH, '//button[text()="검색"]')
                driver.execute_script("arguments[0].click();", search_button)
                time.sleep(0.5)

                # $x('//tr[not(.//div[contains(@class, "header-col")])]')
                claim_tr_list = driver.find_elements(By.XPATH, f'//tr[not(.//div[contains(@class, "header-col")])]')

                claim_tr: webdriver.Chrome._web_element_cls
                for tr_index, claim_tr in enumerate(claim_tr_list):
                    product_dto = ProductDto()

                    order_number = claim_number
                    product_dto.order_number = order_number

                    # 첫번째 행인 경우에만 위치가 다름 (묶음번호가 표기되어있어서)
                    # $x('//tr[not(.//div[contains(@class, "header-col")])]/td[14]')
                    # $x('//tr[not(.//div[contains(@class, "header-col")])]/td[13]')
                    if tr_index == 0:
                        product_name = claim_tr.find_element(By.XPATH, f"./td[14]").get_attribute("textContent").strip()
                    else:
                        product_name = claim_tr.find_element(By.XPATH, f"./td[13]").get_attribute("textContent").strip()

                    product_name = re.sub(r"\[.*?\]", "", product_name)
                    product_name = product_name.strip()
                    product_dto.product_name = product_name

                    # 옵션 td
                    # $x('//td[.//div/span[contains(text(), "- ")]]')
                    product_option = (
                        claim_tr.find_element(By.XPATH, f'./td[.//div/span[contains(text(), "- ")]]')
                        .get_attribute("textContent")
                        .strip()
                    )

                    product_option = product_option[product_option.find("- ") + 2 :]
                    product_option = product_option.replace(": ", "=")
                    product_option = product_option.replace(" - ", ", ")
                    product_dto.product_option = product_option

                    if tr_index == 0:
                        product_qty = claim_tr.find_element(By.XPATH, f"./td[22]").get_attribute("textContent").strip()
                    else:
                        product_qty = claim_tr.find_element(By.XPATH, f"./td[21]").get_attribute("textContent").strip()
                    product_dto.product_qty = product_qty

                    if tr_index == 0:
                        product_recv_name = (
                            claim_tr.find_element(By.XPATH, f"./td[29]").get_attribute("textContent").strip()
                        )
                    else:
                        product_recv_name = (
                            claim_tr.find_element(By.XPATH, f"./td[28]").get_attribute("textContent").strip()
                        )
                    product_dto.product_recv_name = product_recv_name

                    if tr_index == 0:
                        product_recv_tel = (
                            claim_tr.find_element(By.XPATH, f"./td[30]").get_attribute("textContent").strip()
                        )
                    else:
                        product_recv_tel = (
                            claim_tr.find_element(By.XPATH, f"./td[29]").get_attribute("textContent").strip()
                        )
                    product_dto.product_recv_tel = product_recv_tel

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

    def check_order_cancel_number_from_ezadmin(self, order_dict: dict):
        driver = self.driver

        is_checked = False

        order_number = order_dict["주문번호"]
        product_name = order_dict["상품명"]
        product_option: str = order_dict["상품옵션"]
        product_qty = order_dict["수량"]
        product_recv_name = order_dict["수령자명"]
        product_recv_tel = order_dict["수령자연락처"]

        try:
            driver.switch_to.window(self.cs_screen_tab)

            input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "전화번호")]/input')

            input_order_number.clear()

            input_order_number.send_keys(product_recv_tel)

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
                        if not (product_name in search_product_name):
                            continue

                        search_product_option = (
                            driver.find_element(By.XPATH, '//td[@id="di_shop_options"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if not (product_option in search_product_option):
                            continue

                        search_product_qty = (
                            driver.find_element(By.XPATH, '//td[@id="di_order_qty"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if not (product_qty in search_product_qty):
                            continue

                        # search_order_number = (
                        #     driver.find_element(By.XPATH, '//td[@id="di_order_id"]')
                        #     .get_attribute("textContent")
                        #     .strip()
                        # )
                        # if not (order_number in search_order_number):
                        #     continue

                        search_order_number = (
                            driver.find_element(By.XPATH, '//td[@id="di_order_id_seq"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if not (order_number in search_order_number):
                            continue

                        search_product_recv_name = (
                            driver.find_element(By.XPATH, '//td[@id="di_recv_name"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if not (product_recv_name in search_product_recv_name):
                            continue

                        product_cs_state = (
                            driver.find_element(By.XPATH, '//td[@id="di_product_cs"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        if not "배송전 주문취소" in product_cs_state:
                            is_checked = False
                            self.log_msg.emit(f"{self.shop_name} {order_dict}: 배송전 주문취소 상태가 아닙니다.")
                            raise Exception(f"{self.shop_name} {order_dict}: 배송전 주문취소 상태가 아닙니다.")

                        print(
                            f"{search_order_number}: {search_product_name}, {search_product_option}, {search_product_qty}, {product_cs_state}"
                        )

                        is_checked = True

                        if is_checked:
                            return

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
        print(f"Zigzag: work_start")

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
