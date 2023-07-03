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
            driver.get("https://spc.tmon.co.kr/claim/cancel")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h2[text()="출고전 취소 관리"]')))
            time.sleep(0.2)

            # 검색 클릭
            btn_srch = driver.find_element(By.XPATH, '//span[@id="btn_srch" and text()="검색"]')
            driver.execute_script("arguments[0].click();", btn_srch)
            time.sleep(3)

            # 묶음클레임번호 목록
            # $x('//tr[contains(@class, "_xp")]/td[2]/a')
            claim_number_link_list = driver.find_elements(By.XPATH, '//tr[contains(@class, "_xp")]/td[2]/a')

            claim_number_list = []
            for claim_number_link in claim_number_link_list:
                claim_number = claim_number_link.get_attribute("textContent")
                claim_number_list.append(claim_number)

            claim_data = []
            for claim_number in claim_number_list:
                print(f"묶음클레임번호: {claim_number}")

                # 묶음클레임번호 클릭
                try:
                    claim_link = driver.find_element(By.XPATH, f'//a[text()="{claim_number}"]')
                    driver.execute_script("arguments[0].click();", claim_link)
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

                    # $x('//table[@summary="요청정보"]/tbody/tr')
                    claim_product_tr_list = driver.find_elements(By.XPATH, '//table[@summary="요청정보"]/tbody/tr')

                    for claim_product_tr in claim_product_tr_list:
                        product_dto = ProductDto()

                        # 주문번호
                        order_number = driver.find_element(
                            By.XPATH, '//th[text()="주문번호"]/following-sibling::td/strong'
                        ).get_attribute("textContent")
                        product_dto.order_number = order_number

                        # 딜명
                        product_name = claim_product_tr.find_element(By.XPATH, "./td[1]").get_attribute("textContent")
                        product_dto.product_name = product_name

                        # 옵션명
                        product_option = claim_product_tr.find_element(By.XPATH, "./td[2]").get_attribute("textContent")
                        product_dto.product_option = product_option

                        # 수량
                        product_qty = claim_product_tr.find_element(By.XPATH, './td[3][@class="amount"]').get_attribute(
                            "textContent"
                        )
                        product_qty = re.sub(r"[^0-9]", "", product_qty)
                        product_dto.product_qty = product_qty

                        # 수취인명
                        product_recv_name = driver.find_element(
                            By.XPATH, '//div[contains(text()[2], "수취인명")]//text()[2]/following-sibling::em'
                        ).get_attribute("textContent")
                        product_dto.product_recv_name = product_recv_name

                        # 주문자 연락처
                        product_recv_tel = driver.find_element(
                            By.XPATH, '//th[text()="주문자 연락처"]/following-sibling::td/strong'
                        ).get_attribute("textContent")
                        product_dto.product_recv_tel = product_recv_tel

                        product_dto.to_print()

                        claim_data.append({"claim_number": claim_number, "order_number_list": product_dto.get_dict()})

                except Exception as e:
                    print(str(e))
                    print(f"{claim_number} 정보를 수집하지 못했습니다.")
                    continue

                finally:
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
                    raise Exception(f"{self.shop_name} {order}: 배송전 주문취소 상태가 아닙니다.")

            finally:
                driver.refresh()
                driver.switch_to.window(self.shop_screen_tab)

        try:
            driver.get("https://spc.tmon.co.kr/claim/cancel")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h2[text()="출고전 취소 관리"]')))
            time.sleep(0.2)

            # 검색 클릭
            btn_srch = driver.find_element(By.XPATH, '//span[@id="btn_srch" and text()="검색"]')
            driver.execute_script("arguments[0].click();", btn_srch)
            time.sleep(3)

            # 취소 품목 체크박스
            # $x('//tr[./td[contains(@style, "underline") and contains(text(), "442530851")]]//img[contains(@onclick, "cancelBubble")]')
            order_cancel_target_checkbox = driver.find_element(
                By.XPATH,
                f'//tr[.//a[text()="{claim_number}"]]//img[contains(@onclick, "cancelBubble")]',
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(1)

            # 취소처리 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//a[@class="btn_cancel_proc" and text()="취소처리"]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # 해당 요청을 처리할 수 없습니다. alert
            alert_msg = ""
            try:
                WebDriverWait(driver, 1).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert_msg = alert.text
            except Exception as e:
                print(f"no alert")

            print(f"{alert_msg}")

            if ("해당 요청을 처리할 수 없습니다" in alert_msg) or ("처리할 요청건을 선택해 주세요" in alert_msg):
                alert.accept()
                self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")
                raise Exception(f"{self.shop_name} {order}: {alert_msg}")

            elif alert_msg != "":
                alert.accept()
                self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")
                raise Exception(f"{self.shop_name} {order}: {alert_msg}")

            # modal
            try:
                ticketmonster_order_cancel_button = driver.find_element(
                    By.XPATH, '//div[@class="spc_layer claim"][.//h3[text()="취소처리"]]//button[text()="확인"]'
                )
                driver.execute_script("arguments[0].click();", ticketmonster_order_cancel_button)
                time.sleep(0.5)

                # 요청한 1건의 처리가 완료되었습니다.
                # $x('//*[contains(text(), "완료되었습니다")]')
                try:
                    driver.implicitly_wait(20)
                    success_message = driver.find_element(By.XPATH, '//p[@class="message"]').get_attribute(
                        "textContent"
                    )
                    print(success_message)

                    if not "완료되었습니다" in success_message:
                        self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                        raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                except Exception as e:
                    print(str(e))
                    self.log_msg.emit(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")
                    raise Exception(f"{self.shop_name} {order}: 취소 승인 메시지를 찾지 못했습니다.")

                finally:
                    driver.implicitly_wait(self.default_wait)

                self.log_msg.emit(f"{self.shop_name} {order}: 취소 완료")

            except Exception as e:
                print(str(e))
                if self.shop_name in str(e):
                    raise Exception(str(e))

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
