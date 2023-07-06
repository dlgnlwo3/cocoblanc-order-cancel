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


class Naver:
    def __init__(self, log_msg, driver, cs_screen_tab, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.cs_screen_tab: str = cs_screen_tab
        self.dict_account: dict = dict_account

        self.shop_name = "네이버"
        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]

        self.default_wait = 10

    def login(self):
        driver = self.driver

        try:
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//a[@id="loinid"]')))
            time.sleep(0.5)

        except Exception as e:
            print("로그인 화면이 아닙니다.")

        try:
            driver.implicitly_wait(1)

            pyperclip.copy(self.login_id)
            id_input = driver.find_element(By.XPATH, '//input[@id="id"]')
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(Keys.CONTROL, "v")

            pyperclip.copy(self.login_pw)
            pw_input = driver.find_element(By.XPATH, '//input[@id="pw"]')
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(Keys.CONTROL, "v")

            login_button = driver.find_element(By.XPATH, '//button[@id="log.login"]')
            time.sleep(0.2)
            login_button.click()
            time.sleep(0.2)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            driver.implicitly_wait(5)
            main_page_logo = driver.find_element(By.XPATH, '//a[./span[contains(text(), "네이버페이센터")]]')

        except Exception as e:
            print(str(e))
            self.log_msg.emit(f"{self.shop_name} 로그인 실패")
            raise Exception(f"{self.shop_name} 로그인 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            # 새 창 열림
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            naver_popup_tab = other_tabs[0]

            driver.switch_to.window(naver_popup_tab)
            time.sleep(0.5)

            driver.close()

        except Exception as e:
            print(str(e))

        finally:
            driver.switch_to.window(self.shop_screen_tab)

    def get_order_list(self):
        driver = self.driver

        order_list = []
        try:
            driver.get("https://admin.pay.naver.com/o/v3/claim/cancel?summaryInfoType=CANCEL_REQUEST_C1")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "취소관리")]'))
            )
            time.sleep(0.5)

            # 묶음클레임번호 목록
            # $x('//div[@class="tui-grid-lside-area"]//tr[contains(@class, "tui-grid-row")]/td[@data-column-name="productOrderNo"]/div')
            claim_number_link_list = driver.find_elements(
                By.XPATH,
                '//div[@class="tui-grid-lside-area"]//tr[contains(@class, "tui-grid-row")]/td[@data-column-name="productOrderNo"]/div',
            )

            claim_number_list = [
                claim_number_link.get_attribute("textContent") for claim_number_link in claim_number_link_list
            ]

            claim_data = []
            for claim_number in claim_number_list:
                print(f"클레임번호: {claim_number}")

                # 상품주문정보 조회 창을 발생시킨다
                # https://admin.pay.naver.com/o/v3/manage/order/popup/2023053144020740/productOrderDetail
                try:
                    driver.execute_script(
                        f"window.open('https://admin.pay.naver.com/o/v3/manage/order/popup/{claim_number}/productOrderDetail');"
                    )
                    time.sleep(0.2)

                except Exception as e:
                    print(str(e))
                    print(f"{claim_number} 해당 클레임번호를 찾지 못했습니다.")
                    continue

                # 새 창 열림
                try:
                    other_tabs = [
                        tab
                        for tab in driver.window_handles
                        if tab != self.cs_screen_tab and tab != self.shop_screen_tab
                    ]
                    claim_check_tab = other_tabs[0]
                except Exception as e:
                    print(str(e))
                    print(f"{claim_number} 새 창을 찾지 못했습니다.")
                    continue

                # 새 창으로 이동 후 작업
                try:
                    driver.switch_to.window(claim_check_tab)
                    time.sleep(0.5)

                    driver.implicitly_wait(1)

                    product_dto = ProductDto()

                    # 주문번호 -> 묶음번호
                    order_number = ""
                    try:
                        order_number = driver.find_element(
                            By.XPATH, '//th[text()="주문번호"]/following-sibling::td[1]'
                        ).get_attribute("textContent")
                    except Exception as e:
                        print(str(e))
                    product_dto.order_number = order_number

                    # 상품주문번호 -> 상세주문번호
                    order_detail_number = ""
                    try:
                        order_detail_number = driver.find_element(
                            By.XPATH, '//strong[contains(text(), "상품주문번호")]/following-sibling::span[1]'
                        ).get_attribute("textContent")
                    except Exception as e:
                        print(str(e))
                    product_dto.order_detail_number = order_detail_number

                    # 상품명
                    product_name = ""
                    try:
                        product_name = driver.find_element(
                            By.XPATH, '//th[text()="상품명"]/following-sibling::td[1]'
                        ).get_attribute("textContent")
                    except Exception as e:
                        print(str(e))
                    product_dto.product_name = product_name

                    # 옵션
                    product_option = ""
                    try:
                        product_option = driver.find_element(
                            By.XPATH, '//th[text()="옵션"]/following-sibling::td[1]'
                        ).get_attribute("textContent")
                    except Exception as e:
                        print(str(e))
                    product_dto.product_option = product_option

                    # 주문수량
                    product_qty = ""
                    try:
                        product_qty = driver.find_element(
                            By.XPATH, '//th[text()="주문수량"]/following-sibling::td[1]'
                        ).get_attribute("textContent")
                        # product_qty = re.sub(r"[^0-9]", "", product_qty)
                    except Exception as e:
                        print(str(e))
                    product_dto.product_qty = product_qty

                    # 수취인명
                    product_recv_name = ""
                    try:
                        product_recv_name = driver.find_element(
                            By.XPATH, '//th[text()="수취인명"]/following-sibling::td[1]'
                        ).get_attribute("textContent")
                    except Exception as e:
                        print(str(e))
                    product_dto.product_recv_name = product_recv_name

                    # 연락처1
                    product_recv_tel = ""
                    try:
                        product_recv_tel = driver.find_element(
                            By.XPATH, '//th[text()="연락처1"]/following-sibling::td[1]'
                        ).get_attribute("textContent")
                    except Exception as e:
                        print(str(e))
                    product_dto.product_recv_tel = product_recv_tel

                    product_dto.to_print()

                    claim_data.append(
                        {"claim_number": order_detail_number, "order_number_list": product_dto.get_dict()}
                    )

                except Exception as e:
                    print(str(e))
                    print(f"{claim_number} 정보를 수집하지 못했습니다.")
                    continue

                finally:
                    driver.implicitly_wait(self.default_wait)
                    driver.close()
                    driver.switch_to.window(self.shop_screen_tab)

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
            driver.get("https://admin.pay.naver.com/o/v3/claim/cancel?summaryInfoType=CANCEL_REQUEST_C1")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "취소관리")]'))
            )
            time.sleep(0.5)

            # 취소 품목 행과 동일한 lside에 존재하는 radiobutton
            # $x('//tr[.//a[text()="2023051345322390"]]')
            order_cancel_target_radio_button = driver.find_element(
                By.XPATH, f'//tr[.//a[text()="{claim_number}"]]//input[contains(@type, "radio")]'
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_radio_button)
            time.sleep(0.5)

            # 취소 완료처리 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[./span[text()="취소 완료처리"]]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # 새 창 열림 or alert ['선택된 상품주문건이 없습니다.', '취소 승인 처리가 불가능한 상태입니다. 클레임 처리상태를 확인해 주세요.']
            alert_msg = ""
            try:
                WebDriverWait(driver, 1).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert_msg = alert.text
            except Exception as e:
                print(f"no alert")

            print(f"{alert_msg}")

            if ("선택된 상품주문건이 없습니다" in alert_msg) or ("취소 승인 처리가 불가능한 상태입니다" in alert_msg):
                alert.accept()
                self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")
                raise Exception(f"{self.shop_name} {order}: {alert_msg}")

            elif alert_msg != "":
                alert.accept()
                self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")
                raise Exception(f"{self.shop_name} {order}: {alert_msg}")

            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(order_cancel_tab)
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

                order_cancel_button = driver.find_element(By.XPATH, '//a[./span[text()="저장"]]')
                driver.execute_script("arguments[0].click();", order_cancel_button)
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

                    # # 정상적으로 저장되었습니다. alert
                    # try:
                    #     WebDriverWait(driver, 10).until(EC.alert_is_present())
                    # except Exception as e:
                    #     print(f"no alert")
                    #     self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                    #     raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                    # alert_ok_try(driver)
                    # time.sleep(0.5)

                    # 정상적으로 저장되었습니다. alert
                    alert_msg = ""
                    try:
                        WebDriverWait(driver, 10).until(EC.alert_is_present())
                        alert = driver.switch_to.alert
                        alert_msg = alert.text

                    except Exception as e:
                        print(f"no alert")
                        self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                        raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                    print(f"{alert_msg}")

                    # alert_ok_try(driver)

                    if "정상적으로 저장되었습니다" in alert_msg:
                        alert.accept()

                    elif alert_msg != "":
                        alert.accept()
                        self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")

                    else:
                        self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                        raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                elif alert_msg != "":
                    self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")
                    raise Exception(f"{self.shop_name} {order}: {alert_msg}")

                else:
                    self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                    raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                self.log_msg.emit(f"{self.shop_name} {order}: 취소 완료")

            except Exception as e:
                print(str(e))
                driver.close()
                if self.shop_name in str(e):
                    raise Exception(str(e))

            finally:
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))
            if self.shop_name in str(e):
                raise Exception(str(e))

    def work_start(self):
        driver = self.driver
        print(f"Naver: work_start")

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
