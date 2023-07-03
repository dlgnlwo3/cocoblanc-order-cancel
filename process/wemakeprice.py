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


class Wemakeprice:
    def __init__(self, log_msg, driver, cs_screen_tab, dict_account):
        self.log_msg: SignalInstance = log_msg
        self.driver: webdriver.Chrome = driver
        self.cs_screen_tab: str = cs_screen_tab
        self.dict_account: dict = dict_account

        self.shop_name = "위메프"
        self.login_url = self.dict_account["URL"]
        self.login_domain = self.dict_account["도메인"]
        self.login_id = self.dict_account["ID"]
        self.login_pw = self.dict_account["PW"]

        self.default_wait = 10

    def login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 바로 스토어 화면으로 이동합니다.
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//input[@name="loginid"]')))
            time.sleep(2)

        except Exception as e:
            pass

        try:
            driver.implicitly_wait(1)

            id_input = driver.find_element(By.XPATH, '//input[@name="loginid"]')
            time.sleep(0.2)
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(self.login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@name="loginpassword"]')
            time.sleep(0.2)
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
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
                EC.visibility_of_element_located((By.XPATH, '//strong[contains(text(), "취소신청")]'))
            )
            time.sleep(0.5)

        except Exception as e:
            self.log_msg.emit(f"{self.shop_name} 로그인 실패")
            print(e)
            raise Exception(f"{self.shop_name} 로그인 실패")

    def get_order_list(self):
        driver = self.driver

        order_list = []
        try:
            driver.get("https://wpartner.wemakeprice.com/claim/cancelMain")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h2[text()="취소관리"]')))
            time.sleep(0.2)

            # 300개씩 보기
            schLimitCnt_select = Select(driver.find_element(By.XPATH, '//select[@id="schLimitCnt"]'))
            schLimitCnt_select.select_by_visible_text("300개")
            time.sleep(3)

            # 클레임번호 목록
            # $x('//div[@id="claimCancelListGrid"]//tr[contains(@class, "dhx_web")]/td[@rowspan][not(./img)][1]')
            claim_number_list = driver.find_elements(
                By.XPATH,
                '//div[@id="claimCancelListGrid"]//tr[contains(@class, "dhx_web")]/td[@rowspan][not(./img)][1]',
            )

            claim_data = []
            for claim_number in claim_number_list:
                claim_number = claim_number.get_attribute("textContent")

                # $x('//tr[./td[text()="44998480"][not(@title)]]')
                claim_tr_list = driver.find_elements(By.XPATH, f'//tr[./td[text()="{claim_number}"][not(@title)]]')

                claim_tr: webdriver.Chrome._web_element_cls
                for claim_tr in claim_tr_list:
                    product_dto = ProductDto()

                    order_number = (
                        claim_tr.find_element(
                            By.XPATH,
                            f'./td[contains(@style, "underline")][1]',
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.order_number = order_number

                    product_name = (
                        claim_tr.find_element(
                            By.XPATH,
                            f'./td[contains(@style, "underline")][2]/following-sibling::td[1]',
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.product_name = product_name

                    product_option = (
                        claim_tr.find_element(
                            By.XPATH,
                            f'./td[contains(@style, "underline")][2]/following-sibling::td[2]',
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.product_option = product_option

                    product_qty = (
                        claim_tr.find_element(
                            By.XPATH,
                            f'./td[contains(@style, "underline")][2]/following-sibling::td[3]',
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.product_qty = product_qty

                    product_recv_name = (
                        claim_tr.find_element(
                            By.XPATH,
                            f'./td[contains(@style, "underline")][2]/following-sibling::td[@align="center" and @valign="middle"][3]',
                        )
                        .get_attribute("textContent")
                        .strip()
                    )
                    product_dto.product_recv_name = product_recv_name

                    product_recv_tel = (
                        claim_tr.find_element(
                            By.XPATH,
                            f'./td[contains(@style, "underline")][2]/following-sibling::td[@align="center" and @valign="middle"][4]',
                        )
                        .get_attribute("textContent")
                        .strip()
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
            driver.get("https://wpartner.wemakeprice.com/claim/cancelMain")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h2[text()="취소관리"]')))
            time.sleep(0.2)

            # 300개씩 보기
            schLimitCnt_select = Select(driver.find_element(By.XPATH, '//select[@id="schLimitCnt"]'))
            schLimitCnt_select.select_by_visible_text("300개")
            time.sleep(3)

            # 취소 품목 체크박스
            # $x('//tr[./td[text()="44807047"]]//img[contains(@onclick, "cancelBubble")]')
            order_cancel_target = driver.find_element(
                By.XPATH, f'//tr[./td[text()="{claim_number}"]]//img[contains(@onclick, "cancelBubble")]'
            )
            driver.execute_script("arguments[0].click();", order_cancel_target)
            time.sleep(1)

            # 취소승인 버튼
            approveBtn = driver.find_element(By.XPATH, '//button[@id="approveBtn" and text()="취소승인"]')
            driver.execute_script("arguments[0].click();", approveBtn)
            time.sleep(0.5)

            # 취소처리가 가능한 건이 없습니다. 클레임 상태를 확인해 주세요. alert
            alert_msg = ""
            try:
                WebDriverWait(driver, 1).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert_msg = alert.text
            except Exception as e:
                print(f"no alert")
                pass

            print(f"{alert_msg}")

            if alert_msg != "":
                alert.accept()
                self.log_msg.emit(f"{self.shop_name} {order}: {alert_msg}")
                raise Exception(f"{self.shop_name} {order}: {alert_msg}")

            # 새 창 열림
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(order_cancel_tab)
                time.sleep(1)

                order_cancel_button = driver.find_element(By.XPATH, '//button[@id="approveBtn"]')
                driver.execute_script("arguments[0].click();", order_cancel_button)
                time.sleep(0.5)

                # 취소승인 하시겠습니까? alert
                alert_msg = ""
                try:
                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    alert = driver.switch_to.alert
                    alert_msg = alert.text
                except Exception as e:
                    print(f"no alert")
                    pass

                print(f"{alert_msg}")

                if "취소승인 하시겠습니까" in alert_msg:
                    alert.accept()
                    time.sleep(0.5)

                    # 취소승인
                    try:
                        success_count = driver.find_element(By.XPATH, '//strong[@id="returnSuccessCnt"]').get_attribute(
                            "textContent"
                        )

                        if success_count == 0:
                            self.log_msg.emit("성공 건수가 0입니다.")
                            raise Exception("성공 건수가 0입니다.")

                    except Exception as e:
                        print(str(e))
                        self.log_msg.emit(f"{self.shop_name} {order}: {str(e)}")
                        raise Exception(f"{self.shop_name} {order}: {str(e)}")

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
        print(f"Wemakeprice: work_start")

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
