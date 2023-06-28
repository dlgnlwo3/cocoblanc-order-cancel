if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from selenium import webdriver
from dtos.gui_dto import GUIDto

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


class CocoblancOrderCancelProcess:
    def __init__(self):
        self.default_wait = 10
        # open_browser()
        self.driver: webdriver.Chrome = get_chrome_driver(is_headless=False, is_secret=False)
        # self.driver: webdriver.Chrome = get_chrome_driver_new(is_headless=False, is_secret=False)
        self.driver.implicitly_wait(self.default_wait)
        try:
            self.driver.maximize_window()
        except Exception as e:
            print(str(e))

    def setGuiDto(self, guiDto: GUIDto):
        self.guiDto = guiDto

    def setLogger(self, log_msg):
        self.log_msg = log_msg

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

    def ezadmin_login(self):
        driver = self.driver
        self.driver.get(self.dict_accounts["이지어드민"]["URL"])
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//body[@class="ezadmin-main-body"]'))
        )
        time.sleep(0.2)

        login_domain = self.dict_accounts["이지어드민"]["도메인"]
        login_id = self.dict_accounts["이지어드민"]["ID"]
        login_pw = self.dict_accounts["이지어드민"]["PW"]

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
            domain_input.send_keys(login_domain)
            time.sleep(0.2)

            id_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-id"]')
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)
            time.sleep(0.2)

            pwd_input = driver.find_element(By.CSS_SELECTOR, 'input[id="login-pwd"]')
            pwd_input.clear()
            time.sleep(0.2)
            pwd_input.send_keys(login_pw)
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

        self.cs_screen_tab = driver.window_handles[0]

        time.sleep(1)

    # 로그인
    def shop_login(self, account):
        if account == "카카오톡스토어":
            self.kakao_shopping_login()
        elif account == "위메프":
            self.wemakeprice_login()
        elif account == "티몬":
            self.ticketmonster_login()
        elif account == "지그재그":
            self.zigzag_login()
        elif account == "브리치":
            self.bflow_login()
        elif account == "쿠팡":
            self.coupang_login()
        elif account == "11번가":
            self.eleven_street_login()
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

    def bflow_login(self):
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

            login_id = self.dict_accounts["브리치"]["ID"]
            login_pw = self.dict_accounts["브리치"]["PW"]

            id_input = driver.find_element(
                By.XPATH, '//div[@class="login-area"]//input[@class="login-input"][@placeholder="이메일"]'
            )
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(
                By.XPATH, '//div[@class="login-area"]//input[@class="login-input"][@placeholder="비밀번호"]'
            )
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(login_pw)

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
            self.log_msg.emit("지그재그 로그인 실패")
            print(str(e))
            raise Exception("지그재그 로그인 실패")

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

    def eleven_street_login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 바로 스토어 선택 화면으로 이동합니다.
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '//h1[contains(text(), "셀러오피스")]'))
            )
            time.sleep(2)

        except Exception as e:
            pass

        try:
            driver.implicitly_wait(1)

            login_id = self.dict_accounts["11번가"]["ID"]
            login_pw = self.dict_accounts["11번가"]["PW"]

            id_input = driver.find_element(By.XPATH, '//input[@id="loginName"]')
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@id="passWord"]')
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(login_pw)

            login_button = driver.find_element(By.XPATH, '//button[@value="로그인"]')
            time.sleep(0.2)
            login_button.click()
            time.sleep(0.2)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            main_page_link = driver.find_element(By.XPATH, '//h1[./a[contains(text(), "Seller Office")]]')

        except Exception as e:
            self.log_msg.emit("11번가 로그인 실패")
            print(e)
            raise Exception("11번가 로그인 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

    def zigzag_login(self):
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

            login_id = self.dict_accounts["지그재그"]["ID"]
            login_pw = self.dict_accounts["지그재그"]["PW"]

            id_input = driver.find_element(By.XPATH, '//input[contains(@placeholder, "이메일")]')
            id_input.click()
            time.sleep(0.2)
            id_input.send_keys(Keys.LEFT_CONTROL, "a", Keys.BACK_SPACE)
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(By.XPATH, '//input[contains(@placeholder, "비밀번호")]')
            pw_input.click()
            time.sleep(0.2)
            pw_input.send_keys(Keys.LEFT_CONTROL, "a", Keys.BACK_SPACE)
            time.sleep(0.2)
            pw_input.send_keys(login_pw)

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

    def wemakeprice_login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 바로 스토어 화면으로 이동합니다.
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//input[@name="loginid"]')))
            time.sleep(2)

        except Exception as e:
            pass

        try:
            driver.implicitly_wait(1)

            login_id = self.dict_accounts["위메프"]["ID"]
            login_pw = self.dict_accounts["위메프"]["PW"]

            id_input = driver.find_element(By.XPATH, '//input[@name="loginid"]')
            time.sleep(0.2)
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@name="loginpassword"]')
            time.sleep(0.2)
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(login_pw)

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
            self.log_msg.emit(f"위메프 로그인 실패")
            print(e)
            raise Exception("위메프 로그인 실패")

    def coupang_login(self):
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

            login_id = self.dict_accounts["쿠팡"]["ID"]
            login_pw = self.dict_accounts["쿠팡"]["PW"]

            id_input = driver.find_element(By.XPATH, '//input[@id="username"]')
            time.sleep(0.2)
            id_input.click()
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@id="password"]')
            time.sleep(0.2)
            pw_input.click()
            time.sleep(0.2)
            pw_input.clear()
            time.sleep(0.2)
            pw_input.send_keys(login_pw)

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
            self.log_msg.emit(f"쿠팡 로그인 실패")
            raise Exception("쿠팡 로그인 실패")

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

    def kakao_shopping_login(self):
        driver = self.driver

        try:
            # 이전 로그인 세션이 남아있을 경우 해당 web element가 존재하지 않습니다.
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '//button[@type="submit"][contains(text(), "로그인")]'))
            )
            time.sleep(1)

        except Exception as e:
            print(f"카카오쇼핑 로그인 화면이 아닙니다.")

        try:
            driver.implicitly_wait(1)

            login_id = self.dict_accounts["카카오톡스토어"]["ID"]
            login_pw = self.dict_accounts["카카오톡스토어"]["PW"]

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
            print("카카오쇼핑 로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//h1[./a[./img[@alt="톡스토어 판매자센터"]]]'))
            )
            time.sleep(0.5)

        except Exception as e:
            self.log_msg.emit(f"카카오톡스토어 로그인 실패")
            raise Exception("카카오톡스토어 로그인 실패")

        # 각종 팝업창 닫기
        try:
            popup_close_button = driver.find_element(
                By.XPATH, '//div[@class="popup-foot"]//button[./span[contains(text(), "닫기")]]'
            )
            driver.execute_script("arguments[0].click();", popup_close_button)
            time.sleep(0.2)

        except Exception as e:
            print("popup not found")

    # 취소요청 확인
    def get_shop_order_cancel_list(self, account):
        if account == "카카오톡스토어":
            order_cancel_list = self.get_kakao_shopping_order_cancel_list()
        elif account == "위메프":
            order_cancel_list = self.get_wemakeprice_order_cancel_list()
        elif account == "티몬":
            order_cancel_list = self.get_ticketmonster_order_cancel_list()
        elif account == "지그재그":
            order_cancel_list = self.get_zigzag_order_cancel_list()
        elif account == "브리치":
            order_cancel_list = self.get_bflow_order_cancel_list()
        elif account == "쿠팡":
            order_cancel_list = self.get_coupang_order_cancel_list()
        elif account == "11번가":
            order_cancel_list = self.get_eleven_street_order_cancel_list()
        elif account == "네이버":
            order_cancel_list = self.get_naver_order_cancel_list()

        self.log_msg.emit(f"{account}: {len(order_cancel_list)}개의 주문번호(묶음번호)를 발견했습니다.")

        return order_cancel_list

    def get_kakao_shopping_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
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
                return order_cancel_list

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
                claim_number = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_1_")]',
                ).get_attribute("textContent")
                order_number = driver.find_element(
                    By.XPATH,
                    '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding")) and contains(@class, "selected")]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_3_")]',
                ).get_attribute("textContent")
                claim_data.append({"claim_number": claim_number, "order_number_list": order_number})
                send_keys_to_driver(driver, Keys.ARROW_DOWN)

            time.sleep(0.2)

            result = defaultdict(list)

            for item in claim_data:
                result[item["claim_number"]].append(item["order_number_list"])

            result_dict = {claim_number: order_number_list for claim_number, order_number_list in result.items()}

            order_cancel_list = [
                {"claim_number": claim_number, "order_number_list": order_number_list}
                for claim_number, order_number_list in result_dict.items()
            ]

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

        return order_cancel_list

    def get_wemakeprice_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
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

            for claim_number in claim_number_list:
                claim_number = claim_number.get_attribute("textContent")

                # 클레임번호가 포함되어있는 주문번호 목록
                # $x('//tr[./td[text()="44807047"][not(@title)]]/td[contains(@style, "underline")][1]')
                order_number_list = driver.find_elements(
                    By.XPATH,
                    f'//tr[./td[text()="{claim_number}"][not(@title)]]/td[contains(@style, "underline")][1]',
                )

                claim_order_number_list = []
                for order_number in order_number_list:
                    order_number = order_number.get_attribute("textContent")
                    if order_number.isdigit():
                        claim_order_number_list.append(order_number)
                    else:
                        print(f"{order_number}는 숫자가 아닙니다.")

                order_cancel_list.append({"claim_number": claim_number, "order_number_list": claim_order_number_list})

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

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

    def get_zigzag_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
        try:
            driver.get("https://partners.kakaostyle.com/shop/cocoblanc/order_item/list/cancel")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h1[text()="취소관리"]')))
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

            # 주문번호 목록 -> 지그재그의 경우 구매자 연락처
            # $x('//tr[contains(@class, "TableRow")]/td[10]')
            order_number_list = driver.find_elements(By.XPATH, '//tr[contains(@class, "TableRow")]/td[10]')
            phone_number_pattern = r"^01[016789]-\d{3,4}-\d{4}$"
            for order_number in order_number_list:
                order_number = order_number.get_attribute("textContent")
                if re.search(phone_number_pattern, order_number):
                    order_cancel_list.append(order_number)
                else:
                    print(f"{order_number}는 전화번호 양식이 아닙니다.")

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

        return order_cancel_list

    def get_bflow_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
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

            # 주문번호 목록
            order_number_list = driver.find_elements(
                By.XPATH, '//table[contains(@class, "data-table")]//tbody/tr//td[2]'
            )
            for order_number in order_number_list:
                order_number = order_number.get_attribute("textContent")
                order_number = order_number.strip()
                if order_number.isdigit():
                    order_cancel_list.append(order_number)
                else:
                    print(f"{order_number}는 전화번호 양식이 아닙니다.")

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

        return order_cancel_list

    def get_coupang_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []
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

                order_number_list = driver.find_elements(
                    By.XPATH, '//table[.//tr[./th[contains(text(), "출고중지 처리")]]]//tr[not(th)]/td[12]'
                )
                for order_number in reversed(order_number_list):
                    order_number = order_number.get_attribute("textContent")
                    order_number = order_number.strip()
                    if order_number.isdigit():
                        order_cancel_list.insert(0, order_number)
                    else:
                        print(f"{order_number}는 숫자가 아닙니다.")

        except Exception as e:
            print(str(e))

        finally:
            print(f"order_cancel_list: {order_cancel_list}")

        return order_cancel_list

    def get_eleven_street_order_cancel_list(self):
        driver = self.driver

        order_cancel_list = []

        try:
            APIBot = ElevenStreetAPI(self.dict_accounts["11번가"]["API_KEY"])

            driver.get("https://soffice.11st.co.kr/view/6209?preViewCode=D")

            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//iframe[@title="취소관리"]')))
            time.sleep(0.5)

            # 조회 기간은 최대 30일 YYYYMMDDhhmm  'strftime("%Y%m%d%H%M")' 활용
            now = datetime.now()
            startTime = str((now - timedelta(days=30)).strftime("%Y%m%d")) + "0000"
            endTime = str(now.strftime("%Y%m%d")) + "2359"

            # 취소처리 API에는 ordPrdCnSeq, ordNo, ordPrdSeq (클레임번호, 주문번호, 주문순번) 총 세가지 정보가 필요하기 때문에 세개의 정보를 수집해야 함.
            cancelorders = asyncio.run(APIBot.get_cancelorders_from_date(startTime, endTime))

            try:
                if type(cancelorders["ns2:orders"]["ns2:order"]) == dict:
                    api_cancelorder_list = [cancelorders["ns2:orders"]["ns2:order"]]
                else:
                    api_cancelorder_list = cancelorders["ns2:orders"]["ns2:order"]

                for api_cancelorder in api_cancelorder_list:
                    order_cancel_list.insert(
                        0,
                        {
                            "ordPrdCnSeq": api_cancelorder["ordPrdCnSeq"],
                            "ordNo": api_cancelorder["ordNo"],
                            "ordPrdSeq": api_cancelorder["ordPrdSeq"],
                        },
                    )

            except Exception as e:
                print(str(e))

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

                if account == "카카오톡스토어":
                    self.kakao_shopping_order_cancel(account, order_cancel_number)
                elif account == "위메프":
                    self.wemakeprice_order_cancel(account, order_cancel_number)
                elif account == "티몬":
                    self.ticketmonster_order_cancel(account, order_cancel_number)
                elif account == "지그재그":
                    self.zigzag_order_cancel(account, order_cancel_number)
                elif account == "브리치":
                    self.bflow_order_cancel(account, order_cancel_number)
                elif account == "쿠팡":
                    self.coupang_order_cancel(account, order_cancel_number)
                elif account == "11번가":
                    self.eleven_street_order_cancel(account, order_cancel_number)
                elif account == "네이버":
                    self.naver_order_cancel(account, order_cancel_number)

            except Exception as e:
                print(str(e))
                self.log_msg.emit(f"{account} {order_cancel_number}: 작업 실패")
                self.log_msg.emit(f"{str(e)}")
                continue

    def kakao_shopping_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        claim_number = order_cancel_number["claim_number"]
        order_number_list = order_cancel_number["order_number_list"]

        # 주문번호 이지어드민 검증
        for order_number in order_number_list:
            try:
                self.check_order_cancel_number_from_ezadmin(account, order_number)

            except Exception as e:
                print(str(e))
                if order_number in str(e):
                    raise Exception(f"{account} {order_cancel_number}: 배송전 주문취소 상태가 아닙니다.")

        try:
            driver.get("https://store-buy-sell.kakao.com/order/cancelList?orderSummaryCount=CancelRequestToBuyer")

            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "구매자 취소 요청")]'))
            )
            time.sleep(0.2)

            # 주문번호 입력
            input_orderIdList = driver.find_element(By.XPATH, '//input[@name="orderIdList"]')
            input_orderIdList.send_keys(order_number)
            time.sleep(0.2)

            # 검색 클릭
            search_button = driver.find_element(By.XPATH, '//button[@type="submit" and text()="검색"]')
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(1)

            # 취소 품목
            order_cancel_target = driver.find_element(
                By.XPATH,
                f'//button[contains(@onclick, "claim.popOrderDetail") and contains(@onclick, "{order_number}")]',
            )
            driver.execute_script("arguments[0].click();", order_cancel_target)
            time.sleep(1)

            # 새 창 열림
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            kakao_order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(kakao_order_cancel_tab)
                time.sleep(1)

                order_cancel_iframe = driver.find_element(By.XPATH, '//iframe[contains(@src, "omsOrderDetail")]')
                driver.switch_to.frame(order_cancel_iframe)
                time.sleep(0.5)

                kakao_order_cancel_button = driver.find_element(By.XPATH, '//button[contains(text(), "취소승인(환불)")]')
                driver.execute_script("arguments[0].click();", kakao_order_cancel_button)
                time.sleep(0.5)

                # 취소 승인 하시겠습니까? alert
                alert_msg = ""
                try:
                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    alert = driver.switch_to.alert
                    alert_msg = alert.text
                except Exception as e:
                    print(f"no alert")
                    pass

                print(f"{alert_msg}")

                if "취소 승인 하시겠습니까" in alert_msg:
                    alert.accept()

                    # 정상처리 되었습니다. alert
                    try:
                        WebDriverWait(driver, 10).until(EC.alert_is_present())
                    except Exception as e:
                        print(f"no alert")
                        raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

                    alert_ok_try(driver)

                elif alert_msg != "":
                    alert.accept()
                    raise Exception(f"{account} {order_cancel_number}: {alert_msg}")

                else:
                    raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))
                if account in str(e):
                    raise Exception(str(e))

            finally:
                driver.close()
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))

    def wemakeprice_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        claim_number = order_cancel_number["claim_number"]
        order_number_list = order_cancel_number["order_number_list"]

        # 주문번호 이지어드민 검증
        for order_number in order_number_list:
            try:
                self.check_order_cancel_number_from_ezadmin(account, order_number)

            except Exception as e:
                print(str(e))
                if order_number in str(e):
                    raise Exception(f"{account} {order_cancel_number}: 배송전 주문취소 상태가 아닙니다.")

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

            # 새 창 열림 or alert ['선택된 상품주문건이 없습니다.', '취소처리가 가능한 건이 없습니다. 클레임 상태를 확인해 주세요.']
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            wemakeprice_order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(wemakeprice_order_cancel_tab)
                time.sleep(1)

                wemakeprice_order_cancel_button = driver.find_element(By.XPATH, '//button[@id="approveBtn"]')
                driver.execute_script("arguments[0].click();", wemakeprice_order_cancel_button)
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
                            raise Exception("성공 건수가 0입니다.")

                    except Exception as e:
                        print(str(e))
                        raise Exception(f"{account} {order_cancel_number}: {str(e)}")

                elif alert_msg != "":
                    alert.accept()
                    raise Exception(f"{account} {order_cancel_number}: {alert_msg}")

                else:
                    raise Exception(f"{account} {order_cancel_number}: 취소 승인 메시지를 찾지 못했습니다.")

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))
                if account in str(e):
                    raise Exception(str(e))

            finally:
                driver.close()
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))
            else:
                # 취소처리가 가능한 건이 없습니다. 클레임 상태를 확인해 주세요.
                failed_alert_msg = ""
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present())
                    failed_alert = driver.switch_to.alert
                    failed_alert_msg = failed_alert.text
                except Exception as e:
                    print(f"no alert")

                print(f"failed_alert_msg: {failed_alert_msg}")

                if failed_alert_msg != "":
                    raise Exception(f"{account} {order_cancel_number}: {failed_alert_msg}")
                else:
                    raise Exception(f"{account} {order_cancel_number}: 해당 주문이 존재하지 않습니다.")

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

    def zigzag_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

        try:
            driver.get("https://partners.kakaostyle.com/shop/cocoblanc/order_item/list/cancel")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h1[text()="취소관리"]')))
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

            # 취소 품목 체크박스
            # $x('//tr[./td[contains(@style, "underline") and contains(text(), "442530851")]]//img[contains(@onclick, "cancelBubble")]')
            order_cancel_target_checkbox = driver.find_element(
                By.XPATH,
                f'//tr[./td[text()="{order_cancel_number}"]]//th[contains(@class, "checkbox")]//input',
            )
            driver.execute_script("arguments[0].click();", order_cancel_target_checkbox)
            time.sleep(1)

            # 취소완료 버튼 클릭 시 modal창 열림
            btn_cancel_proc = driver.find_element(By.XPATH, '//button[.//span[text()="취소완료"]]')
            driver.execute_script("arguments[0].click();", btn_cancel_proc)
            time.sleep(0.5)

            try:
                zigzag_order_cancel_button = driver.find_element(By.XPATH, '//button[contains(text(), "취소완료")]')
                driver.execute_script("arguments[0].click();", zigzag_order_cancel_button)
                time.sleep(1)

                # 선택하신 1건의 상품주문을 환불처리 하시겠습니까?
                cancel_agree_button = driver.find_element(
                    By.XPATH,
                    '//div[@class="modal-content"][.//div[contains(text(), "환불처리 하시겠습니까?")]]//button[text()="확인"]',
                )
                driver.execute_script("arguments[0].click();", cancel_agree_button)
                time.sleep(5)

                # 1개의 상품주문이 취소 완료 처리 되었습니다.
                cancel_success_message = driver.find_element(
                    By.XPATH, '//div[contains(text(), "취소 완료 처리 되었습니다")]'
                ).get_attribute("textContent")
                print(cancel_success_message)

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))
                if account in str(e):
                    raise Exception(str(e))
                else:
                    raise Exception(f"{account} {order_cancel_number}: 취소 완료 메시지를 찾지 못했습니다.")

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

    def eleven_street_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 11번가 order_cancel_number -> ordPrdCnSeq, ordNo, ordPrdSeq (클레임번호, 주문번호, 주문순번) 세개의 정보가 담겨있음
        # 주문번호 이지어드민 검증
        ordPrdCnSeq = order_cancel_number["ordPrdCnSeq"]
        ordNo = order_cancel_number["ordNo"]
        ordPrdSeq = order_cancel_number["ordPrdSeq"]

        try:
            self.check_order_cancel_number_from_ezadmin(account, ordNo)

        except Exception as e:
            print(str(e))
            if ordNo in str(e):
                raise Exception((str(e)))

        try:
            APIBot = ElevenStreetAPI(self.dict_accounts["11번가"]["API_KEY"])

            driver.get("https://soffice.11st.co.kr/view/6209?preViewCode=D")

            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//iframe[@title="취소관리"]')))
            time.sleep(0.5)

            ResultOrder = asyncio.run(APIBot.cancelreqconf_from_ordInfo(ordPrdCnSeq, ordNo, ordPrdSeq))

            if ResultOrder["ResultOrder"]["result_code"] != "0":
                raise Exception(f"{account} {ordNo}: 취소 승인 메시지를 찾지 못했습니다.")

            self.log_msg.emit(f"{account} {ordNo}: 취소 완료")

        except Exception as e:
            print(str(e))
            if account in str(e):
                raise Exception(str(e))
            else:
                raise Exception(f"{account} {ordNo}: 해당 주문이 존재하지 않습니다.")

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

    def check_order_cancel_number_from_ezadmin(self, account, order_cancel_number):
        driver = self.driver
        try:
            driver.switch_to.window(self.cs_screen_tab)

            if account == "네이버" or account == "지그재그":
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "전화번호")]/input')
            else:
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "주문번호")]/input')

            input_order_number.clear()
            time.sleep(0.2)

            input_order_number.send_keys(order_cancel_number)
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

                        product_cs_state = (
                            driver.find_element(By.XPATH, '//td[@id="di_product_cs"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        product_name = (
                            driver.find_element(By.XPATH, '//td[@id="di_shop_pname"]')
                            .get_attribute("textContent")
                            .strip()
                        )

                        product_option = (
                            driver.find_element(By.XPATH, '//td[@id="di_shop_options"]')
                            .get_attribute("textContent")
                            .strip()
                        )

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

    def check_order_cancel_number_from_ezadmin_for_phone(self, account, order_cancel_number):
        driver = self.driver
        try:
            driver.switch_to.window(self.cs_screen_tab)

            if account == "네이버" or account == "지그재그":
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "전화번호")]/input')
            else:
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "주문번호")]/input')

            input_order_number.clear()
            time.sleep(0.2)

            input_order_number.send_keys(order_cancel_number)
            time.sleep(0.2)

            search_button = driver.find_element(By.XPATH, '//div[@id="search"][text()="검색"]')
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(3)

            # # 위쪽 테이블 -> $x('//table[@id="grid_order"]')
            # try:
            #     order_cancel_number_td = driver.find_element(
            #         By.XPATH, f'//table[@id="grid_order"]//td[@title="{order_cancel_number}"]'
            #     )
            #     driver.execute_script("arguments[0].click();", order_cancel_number_td)
            # except Exception as e:
            #     raise Exception(f"{account} {order_cancel_number}: 이지어드민 검색 결과가 없습니다.")
            # finally:
            #     time.sleep(0.5)

            # 아래쪽 테이블 -> $x('//table[contains(@id, "grid_product")]')
            cs_state_trs = driver.find_elements(
                By.XPATH,
                f'//table[contains(@id, "grid_product")]//td[contains(@title, "list_order_id") and contains(@title, "{order_cancel_number}")]',
            )

            if len(cs_state_trs) == 0:
                raise Exception(f"{account} {order_cancel_number}: 이지어드민 검색 결과가 없습니다.")

            for cs_state_tr in cs_state_trs:
                driver.execute_script("arguments[0].click();", cs_state_tr)
                product_cs_state = driver.find_element(By.XPATH, '//td[@id="di_product_cs"]')
                product_cs_state = product_cs_state.get_attribute("textContent")
                print(f"{account}, {order_cancel_number}, {product_cs_state}")

                if not "배송전 주문취소" in product_cs_state:
                    raise Exception(f"{account} {order_cancel_number}: 배송전 주문취소 상태가 아닙니다.")

            # 아래쪽 테이블의 검색결과 탭을 닫는 버튼 -> $x('//span[contains(@class, "ui-icon-close")]')

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception(str(e))

        finally:
            driver.refresh()
            driver.switch_to.window(self.shop_screen_tab)

    # 전체작업 시작
    def work_start(self):
        driver = self.driver

        print(f"process: work_start")

        try:
            self.dict_accounts = self.get_dict_account()

            # 로그인
            self.ezadmin_login()

            # cs창 진입
            self.switch_to_cs_screen()

            # 쇼핑몰의 수 만큼 작업
            for account in self.dict_accounts:
                try:
                    if account == "이지어드민":
                        continue

                    # # 쇼핑몰 제외 테스트용 코드
                    # if account == "위메프":
                    #     continue

                    # # 쇼핑몰 단일 테스트용 코드
                    # if account != "네이버":
                    #     continue

                    print(account)
                    account_url = self.dict_accounts[account]["URL"]

                    try:
                        driver.execute_script(f"window.open('{account_url}');")
                        time.sleep(1)
                        self.shop_screen_tab = driver.window_handles[1]
                        driver.switch_to.window(self.shop_screen_tab)
                        time.sleep(1)

                        # 쇼핑몰 로그인
                        self.shop_login(account)

                        # 쇼핑몰 취소요청 확인
                        order_cancel_list = self.get_shop_order_cancel_list(account)

                        # 쇼핑몰 취소
                        self.shop_order_cancel(account, order_cancel_list)

                    except Exception as e:
                        print(str(e))
                        print(f"{account} 작업 실패")

                    finally:
                        driver.close()
                        driver.switch_to.window(self.cs_screen_tab)

                        self.log_msg.emit(f"{account}: 작업 종료")

                    time.sleep(1)

                except Exception as e:
                    print(str(e))

        except Exception as e:
            print(str(e))


if __name__ == "__main__":
    # process = CocoblancOrderCancelProcess()
    # process.work_start()
    pass
