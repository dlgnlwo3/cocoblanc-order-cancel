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

from process.ezadmin import Ezadmin
from process.kakaotalk_store import KakaoTalkStore
from process.wemakeprice import Wemakeprice
from process.zigzag import Zigzag
from process.bflow import Bflow
from process.coupang import Coupang
from process.eleven_street import ElevenStreet


class CocoblancOrderCancelProcess:
    def __init__(self):
        self.driver: webdriver.Chrome = get_chrome_driver(is_headless=False, is_secret=False)
        self.default_wait = 10
        self.driver.implicitly_wait(self.default_wait)
        try:
            self.driver.maximize_window()
        except Exception as e:
            print(str(e))

    def setGuiDto(self, guiDto: GUIDto):
        self.guiDto = guiDto

    def setLogger(self, log_msg):
        self.log_msg: SignalInstance = log_msg

    def get_dict_account(self):
        df_accounts = AccountFile(self.guiDto.account_file).df_account
        df_accounts = df_accounts.fillna("")
        dict_accounts = {}
        for index, row in df_accounts.iterrows():
            channel = str(row["채널명"])
            domain = str(row["도메인"])
            account_id = str(row["ID"])
            account_pw = str(row["PW"])
            url = str(row["URL"])
            api_key = str(row["API_KEY"])
            dict_accounts[channel] = {"도메인": domain, "ID": account_id, "PW": account_pw, "URL": url, "API_KEY": api_key}
        return dict_accounts

    # 로그인
    def shop_login(self, account):
        if account == "티몬":
            self.ticketmonster_login()
        elif account == "네이버":
            self.naver_login()

        # self.log_msg.emit(f"{account}: 로그인 성공")

    def naver_login(self):
        driver = self.driver

        try:
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//a[@id="loinid"]')))
            time.sleep(0.5)

        except Exception as e:
            print("로그인 화면이 아닙니다.")

        try:
            driver.implicitly_wait(1)

            login_id = self.dict_accounts["네이버"]["ID"]
            login_pw = self.dict_accounts["네이버"]["PW"]

            pyperclip.copy(login_id)
            id_input = driver.find_element(By.XPATH, '//input[@id="id"]')
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(Keys.CONTROL, "v")

            pyperclip.copy(login_pw)
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
            self.log_msg.emit("네이버 로그인 실패")
            print(str(e))
            raise Exception("네이버 로그인 실패")

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

    def ticketmonster_login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 바로 스토어 선택 화면으로 이동합니다.
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '//h2[contains(text(), "파트너  로그인")]'))
            )
            time.sleep(2)

        except Exception as e:
            pass

        try:
            driver.implicitly_wait(1)

            login_id = self.dict_accounts["티몬"]["ID"]
            login_pw = self.dict_accounts["티몬"]["PW"]

            id_input = driver.find_element(By.XPATH, '//input[@id="form_id"]')
            time.sleep(0.2)
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@id="form_password"]')
            time.sleep(0.2)
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(login_pw)

            login_button = driver.find_element(By.XPATH, '//button[contains(@onclick, "submitLogin()")]')
            login_button.click()
            time.sleep(1)

            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "다음에 변경")]'))
            )
            time.sleep(3)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            change_next_time_button = driver.find_element(By.XPATH, '//button[contains(text(), "다음에 변경")]')
            time.sleep(0.2)
            change_next_time_button.click()
            time.sleep(1)
        except Exception as e:
            print("비밀번호 변경 안내 스킵 실패")

        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//h3[contains(text(), "취소/환불/교환 현황")]'))
            )
            time.sleep(0.5)

        except Exception as e:
            self.log_msg.emit(f"티몬 로그인 실패")
            print(e)
            raise Exception("티몬 로그인 실패")

    # 취소요청 확인
    def get_shop_order_cancel_list(self, account):
        if account == "티몬":
            order_cancel_list = self.get_ticketmonster_order_cancel_list()
        elif account == "네이버":
            order_cancel_list = self.get_naver_order_cancel_list()

        self.log_msg.emit(f"{account}: {len(order_cancel_list)}개의 주문번호(묶음번호)를 발견했습니다.")

        return order_cancel_list

    def get_ticketmonster_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
        try:
            driver.get("https://spc.tmon.co.kr/claim/cancel")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h2[text()="출고전 취소 관리"]')))
            time.sleep(0.2)

            # 검색 클릭
            btn_srch = driver.find_element(By.XPATH, '//span[@id="btn_srch" and text()="검색"]')
            driver.execute_script("arguments[0].click();", btn_srch)
            time.sleep(3)

            # 주문번호 목록
            # $x('//tr[contains(@class, "_xp")]//a[contains(@href, "popupBuyDetailInfo")]')
            order_number_list = driver.find_elements(
                By.XPATH, '//tr[contains(@class, "_xp")]//a[contains(@href, "popupBuyDetailInfo")]'
            )
            for order_number in order_number_list:
                order_number = order_number.get_attribute("textContent")
                if order_number.isdigit():
                    order_cancel_list.append(order_number)
                else:
                    print(f"{order_number}는 숫자가 아닙니다.")

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

        return order_cancel_list

    def get_naver_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
        try:
            driver.get("https://admin.pay.naver.com/o/v3/claim/cancel?summaryInfoType=CANCEL_REQUEST_C1")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "취소관리")]'))
            )
            time.sleep(0.5)

            # 주문번호 목록 -> 네이버의 경우 구매자 연락처
            # $x('//tr/td[@data-column-name="receiverTelNo1"]/div')
            order_number_list = driver.find_elements(By.XPATH, '//tr/td[@data-column-name="receiverTelNo1"]/div')
            phone_number_pattern = r"^01[016789]-\d{3,4}-\d{4}$"
            virtual_number_pattern = r"050\d{1}-\d{3,4}-\d{3,4}"
            for order_number in order_number_list:
                order_number = order_number.get_attribute("textContent")
                if re.search(phone_number_pattern, order_number) or re.search(virtual_number_pattern, order_number):
                    order_cancel_list.append(order_number)
                else:
                    print(f"{order_number}는 전화번호, 안심번호 양식이 아닙니다.")

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

        return order_cancel_list

    def shop_order_cancel(self, account, order_cancel_list):
        for order_cancel_number in order_cancel_list:
            try:
                print(f"{account} {order_cancel_number}")

                if account == "티몬":
                    self.ticketmonster_order_cancel(account, order_cancel_number)
                elif account == "네이버":
                    self.naver_order_cancel(account, order_cancel_number)

            except Exception as e:
                print(str(e))
                self.log_msg.emit(f"{account} {order_cancel_number}: 작업 실패")
                self.log_msg.emit(f"{str(e)}")
                continue

    def ticketmonster_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

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
                f'//tr[.//a[text()="{order_cancel_number}"]]//img[contains(@onclick, "cancelBubble")]',
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(1)

            # 취소처리 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//a[@class="btn_cancel_proc" and text()="취소처리"]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

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
                        raise Exception(f"{account} {order_cancel_number}: 취소 완료 메시지를 찾지 못했습니다.")

                except Exception as e:
                    print(str(e))
                    raise Exception(f"{account} {order_cancel_number}: 취소 완료 메시지를 찾지 못했습니다.")

                finally:
                    driver.implicitly_wait(self.default_wait)

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))
                if account in str(e):
                    raise Exception(str(e))

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{account} {order_cancel_number}: 해당 주문이 존재하지 않습니다.")

    def bflow_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

        try:
            driver.get("https://b-flow.co.kr/order/cancels")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "취소 관리")]'))
            )
            time.sleep(0.2)

            # 로딩바 존재
            # $x('//div[contains(@class, "overlay")]')
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
            time.sleep(0.5)

            # 취소 품목 체크박스
            # '//tr[./td[.//span[contains(text(), "2023061616868779870")]]]//input[@type="checkbox"]'
            order_cancel_target_checkbox = driver.find_element(
                By.XPATH, f'//tr[./td[.//span[contains(text(), "{order_cancel_number}")]]]//input[@type="checkbox"]'
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
                raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

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
                    raise Exception(f"{account} {order_cancel_number}: 취소 성공 메시지를 찾지 못했습니다.")

            elif alert_msg != "":
                alert.dismiss()
                raise Exception(f"{account} {order_cancel_number}: {alert_msg}")

            else:
                raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

            self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{account} {order_cancel_number}: 해당 주문이 존재하지 않습니다.")

    def coupang_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

        try:
            driver.get("https://wing.coupang.com/tenants/sfl-portal/stop-shipment/list")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h3[contains(text(), "출고중지관리")]'))
            )
            time.sleep(0.5)

            # 취소 품목 체크박스
            # '//tr[./td[.//span[contains(text(), "2023061616868779870")]]]//input[@type="checkbox"]'
            order_cancel_target_checkbox = driver.find_element(
                By.XPATH, f'//tr[.//a[contains(text(), "{order_cancel_number}")]]/td//input[@type="checkbox"]'
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(1)

            # 출고중지완료 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[contains(text(), "출고중지완료")]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # modal창: 하기 1건을 출고중지 완료 하시겠습니까? 출고중지완료하시면 환불이 완료됩니다.
            # $x('//div[@data-wuic-partial="widget"][.//span[contains(text(), "출고중지완료하시면 환불이 완료됩니다.")]]')
            try:
                coupang_order_cancel_button = driver.find_element(
                    By.XPATH,
                    '//div[@data-wuic-partial="widget"][.//span[contains(text(), "출고중지완료하시면 환불이 완료됩니다.")]]//div[@class="footer"]/button[contains(text(), "완료")]',
                )
                driver.execute_script("arguments[0].click();", coupang_order_cancel_button)
                time.sleep(2)

                # 별다른 메시지 없이 처리됨

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))
                if account in str(e):
                    raise Exception(str(e))

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{account} {order_cancel_number}: 해당 주문이 존재하지 않습니다.")

    def naver_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

        try:
            driver.get("https://admin.pay.naver.com/o/v3/claim/cancel?summaryInfoType=CANCEL_REQUEST_C1")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//h2[contains(text(), "취소관리")]'))
            )
            time.sleep(0.5)

            # 취소 품목의 행 번호 -> rside
            # $x('//tr[./td[@data-column-name="receiverTelNo1"]/div[text()="010-4134-0899"]]')
            order_cancel_target_row = driver.find_element(
                By.XPATH, f'//tr[./td[@data-column-name="receiverTelNo1"]/div[text()="{order_cancel_number}"]]'
            ).get_attribute("data-row-key")

            # 취소 품목 행과 동일한 lside에 존재하는 radiobutton
            # $x('//div[@class="tui-grid-lside-area"]//tr[@data-row-key="1"]//input[contains(@type, "radio")]')
            order_cancel_target_checkbox = driver.find_element(
                By.XPATH,
                f'//div[@class="tui-grid-lside-area"]//tr[@data-row-key="{order_cancel_target_row}"]//input[contains(@type, "radio")]',
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(0.5)

            # 취소 완료처리 버튼
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[./span[text()="취소 완료처리"]]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            # 새 창 열림 or alert ['선택된 상품주문건이 없습니다.', '취소 승인 처리가 불가능한 상태입니다. 클레임 처리상태를 확인해 주세요.']
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            naver_order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(naver_order_cancel_tab)
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

                naver_order_cancel_button = driver.find_element(By.XPATH, '//a[./span[text()="저장"]]')
                driver.execute_script("arguments[0].click();", naver_order_cancel_button)
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

                    # 정상적으로 저장되었습니다. alert
                    try:
                        WebDriverWait(driver, 10).until(EC.alert_is_present())
                    except Exception as e:
                        print(f"no alert")
                        raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

                    alert_ok_try(driver)
                    time.sleep(0.5)

                elif alert_msg != "":
                    raise Exception(f"{account} {order_cancel_number}: {alert_msg}")

                else:
                    raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))
                driver.close()
                if account in str(e):
                    raise Exception(str(e))

            finally:
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{account} {order_cancel_number}: 해당 주문이 존재하지 않습니다.")

    def check_order_cancel_number_from_ezadmin(self, account, order_cancel_number: dict):
        driver = self.driver
        try:
            driver.switch_to.window(self.cs_screen_tab)

            if account == "네이버" or account == "지그재그":
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "전화번호")]/input')
            else:
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "주문번호")]/input')

            input_order_number.clear()
            time.sleep(0.2)

            input_order_number.send_keys(order_cancel_number["주문번호"])
            time.sleep(0.2)

            search_button = driver.find_element(By.XPATH, '//div[@id="search"][text()="검색"]')
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(3)

            grid_order_trs = driver.find_elements(By.XPATH, '//table[@id="grid_order"]//tr[not(@class="jqgfirstrow")]')

            if len(grid_order_trs) == 0:
                raise Exception(f"{account} {order_cancel_number}: 이지어드민 검색 결과가 없습니다.")

            for grid_order_tr in grid_order_trs:
                try:
                    driver.execute_script("arguments[0].click();", grid_order_tr)
                    time.sleep(0.2)

                    grid_product_trs = driver.find_elements(
                        By.XPATH,
                        f'//table[contains(@id, "grid_product")]//td[contains(@title, "list_order_id") and contains(@title, "{order_cancel_number}")]',
                    )

                    if len(grid_product_trs) == 0:
                        raise Exception(f"{account} {order_cancel_number}: 이지어드민 검색 결과가 없습니다.")

                    for grid_product_tr in grid_product_trs:
                        driver.execute_script("arguments[0].click();", grid_product_tr)
                        time.sleep(0.2)

                        product_name = (
                            driver.find_element(By.XPATH, '//td[@id="di_shop_pname"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if (order_cancel_number["상품명"] in product_name) or (
                            order_cancel_number["상품옵션"] in product_name
                        ):
                            pass

                        product_option = (
                            driver.find_element(By.XPATH, '//td[@id="di_shop_options"]')
                            .get_attribute("textContent")
                            .strip()
                        )
                        if order_cancel_number["상품옵션"] in product_option:
                            pass

                        product_qty = (
                            driver.find_element(By.XPATH, '//td[@id="di_order_qty"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        product_order_id_seq = (
                            driver.find_element(By.XPATH, '//td[@id="di_order_id_seq"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        print(
                            f"{account}, {order_cancel_number}, {product_cs_state}, {product_name}, {product_option}, {product_qty}, {product_order_id_seq}"
                        )

                        product_cs_state = (
                            driver.find_element(By.XPATH, '//td[@id="di_product_cs"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        if not "배송전 주문취소" in product_cs_state:
                            raise Exception(f"{account} {order_cancel_number}: 배송전 주문취소 상태가 아닙니다.")

                except Exception as e:
                    print(str(e))
                    if account in str(e):
                        raise Exception(str(e))

                finally:
                    tab_close_button = driver.find_element(By.XPATH, '//span[contains(@class, "ui-icon-close")]')
                    driver.execute_script("arguments[0].click();", tab_close_button)
                    time.sleep(0.2)

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))

        finally:
            driver.refresh()
            driver.switch_to.window(self.shop_screen_tab)

    # 전체작업 시작
    def work_start(self):
        print(f"CocoblancOrderCancelProcess: work_start")

        try:
            self.dict_accounts = self.get_dict_account()

            # 쇼핑몰의 수 만큼 작업 (계정 엑셀 파일의 행)
            for account in self.dict_accounts:
                try:
                    dict_account = self.dict_accounts[account]

                    if account == "이지어드민":
                        ezadmin = Ezadmin(self.log_msg, self.driver, dict_account)
                        ezadmin.login()
                        self.cs_screen_tab = ezadmin.switch_to_cs_screen()

                    # if account == "카카오톡스토어":
                    #     kakaotalk_store = KakaoTalkStore(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     kakaotalk_store.work_start()

                    # if account == "위메프":
                    #     wemakeprice = Wemakeprice(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     wemakeprice.work_start()

                    # if account == "티몬":
                    #     pass

                    # if account == "지그재그":
                    #     zigzag = Zigzag(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     zigzag.work_start()

                    if account == "브리치":
                        bflow = Bflow(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                        bflow.work_start()

                    # if account == "쿠팡":
                    #     coupang = Coupang(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     coupang.work_start()

                    # if account == "11번가":
                    #     eleven_street = ElevenStreet(self.log_msg, self.driver, self.cs_screen_tab, dict_account)
                    #     eleven_street.work_start()

                    if account == "네이버":
                        pass

                except Exception as e:
                    print(str(e))

        except Exception as e:
            print(str(e))


if __name__ == "__main__":
    # process = CocoblancOrderCancelProcess()
    # process.work_start()
    pass
