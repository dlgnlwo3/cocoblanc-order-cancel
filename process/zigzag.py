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

            claim_number_list = [
                claim_number_tr.get_attribute("textContent") for claim_number_tr in claim_number_tr_list
            ]
            claim_number_list = list(dict.fromkeys(claim_number_list))
            # for claim_number_tr in claim_number_tr_list:
            #     claim_number = claim_number_tr.get_attribute("textContent")
            #     claim_number_list.append(claim_number)

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

                # 검색
                search_button = driver.find_element(By.XPATH, '//button[text()="검색"]')
                driver.execute_script("arguments[0].click();", search_button)
                time.sleep(0.5)

                # $x('//tr[not(.//div[contains(@class, "header-col")])]/td')
                # 일반적으로 36개의 td를 갖고있음
                # 묶음번호에 묶여있다면 35개의 td를 갖고있음
                claim_tr_list = driver.find_elements(By.XPATH, f'//tr[not(.//div[contains(@class, "header-col")])]')

                claim_tr: webdriver.Chrome._web_element_cls
                for tr_index, claim_tr in enumerate(claim_tr_list):
                    product_dto = ProductDto()

                    order_number = claim_number
                    product_dto.order_number = order_number

                    td_len = len(claim_tr.find_elements(By.XPATH, "./td"))

                    if td_len == 36:
                        product_name = claim_tr.find_element(By.XPATH, f"./td[14]").get_attribute("textContent").strip()
                    else:
                        product_name = claim_tr.find_element(By.XPATH, f"./td[13]").get_attribute("textContent").strip()

                    product_name = re.sub(r"\[.*?\]", "", product_name)
                    product_name = product_name.strip()
                    product_dto.product_name = product_name

                    # 옵션 td
                    # $x('//tr[not(.//div[contains(@class, "header-col")])]/td[18]')
                    if td_len == 36:
                        product_option = (
                            claim_tr.find_element(By.XPATH, f"./td[18]").get_attribute("textContent").strip()
                        )
                    else:
                        product_option = (
                            claim_tr.find_element(By.XPATH, f"./td[17]").get_attribute("textContent").strip()
                        )

                    product_option = product_option.replace(": ", "=")
                    option_slice = product_option.find(" ")
                    product_option = product_option[:option_slice] + ", " + product_option[option_slice + 1 :]
                    product_dto.product_option = product_option

                    if td_len == 36:
                        product_qty = claim_tr.find_element(By.XPATH, f"./td[22]").get_attribute("textContent").strip()
                    else:
                        product_qty = claim_tr.find_element(By.XPATH, f"./td[21]").get_attribute("textContent").strip()
                    product_dto.product_qty = product_qty

                    if td_len == 36:
                        product_recv_name = (
                            claim_tr.find_element(By.XPATH, f"./td[29]").get_attribute("textContent").strip()
                        )
                    else:
                        product_recv_name = (
                            claim_tr.find_element(By.XPATH, f"./td[28]").get_attribute("textContent").strip()
                        )
                    product_dto.product_recv_name = product_recv_name

                    if td_len == 36:
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
            number_input = driver.find_element(By.XPATH, '//label[./span[text()="상세조건"]]/following-sibling::div//input')
            number_input.clear()
            number_input.send_keys(claim_number)
            time.sleep(0.2)

            # 검색
            search_button = driver.find_element(By.XPATH, '//button[text()="검색"]')
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(0.5)

            # 취소 품목 체크박스
            # $x('//tr[.//button[contains(text(), "138752587359305632")]]//th[contains(@class, "checkbox")]//input')
            # order_cancel_target_checkbox = driver.find_element(
            #     By.XPATH,
            #     f'//tr[.//button[contains(text(), "{claim_number}")]]//th[contains(@class, "checkbox")]//input',
            # )

            # 전체 체크 체크박스
            order_cancel_target_checkbox = driver.find_element(By.XPATH, f'//thead//input[@type="checkbox"]')
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(1)

            # 취소완료 버튼 클릭 시 modal창 열림
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[.//span[text()="취소완료"]]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # 체크박스로 1개의 상품만 체크된 경우
            try:
                driver.implicitly_wait(1)
                zigzag_order_cancel_button = driver.find_element(By.XPATH, '//button[contains(text(), "취소완료")]')
                driver.execute_script("arguments[0].click();", zigzag_order_cancel_button)
                time.sleep(1)
            except Exception as e:
                print(f"2개 이상의 결과")
            finally:
                driver.implicitly_wait(self.default_wait)

            try:
                # 선택하신 1건의 상품주문을 환불처리 하시겠습니까? or 선택하신 2개의 상품 주문을 취소처리 하시겠습니까?
                cancel_agree_button = driver.find_element(
                    By.XPATH,
                    '//div[@class="modal-content"][.//div[contains(text(), "하시겠습니까?")]]//button[text()="확인"]',
                )
                driver.execute_script("arguments[0].click();", cancel_agree_button)
                time.sleep(5)

                # 1개의 상품주문이 취소 완료 처리 되었습니다. or 2개의 상품주문이 취소 완료 처리 되었습니다.
                try:
                    cancel_success_message = driver.find_element(
                        By.XPATH, '//div[contains(text(), "취소 완료 처리 되었습니다")]'
                    ).get_attribute("textContent")
                    print(cancel_success_message)
                except Exception as e:
                    self.log_msg.emit(f"{self.shop_name} {order}: 취소 완료 메시지를 찾지 못했습니다.")
                    raise Exception(f"{self.shop_name} {order}: 취소 완료 메시지를 찾지 못했습니다.")

                self.log_msg.emit(f"{self.shop_name} {order}: 취소 완료")

            except Exception as e:
                print(str(e))
                if self.shop_name in str(e):
                    raise Exception(str(e))
                else:
                    raise Exception(f"{self.shop_name} {order}: 취소 완료 메시지를 찾지 못했습니다.")

        except Exception as e:
            print(str(e))
            if self.shop_name in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{self.shop_name} {order}: 해당 주문이 존재하지 않습니다.")

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
