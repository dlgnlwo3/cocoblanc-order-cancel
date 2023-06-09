if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from selenium import webdriver
from dtos.gui_dto import GUIDto

from common.utils import global_log_append
from common.chrome import open_browser, get_chrome_driver, get_chrome_driver_new
from common.selenium_activities import close_new_tabs, alert_ok_try
from common.account_file import AccountFile


from enums.store_column_enum import CommonStoreEnum, Cafe24Enum, ElevenStreetEnum
from enums.store_name_enum import StoreNameEnum

from features.convert_store_name import StoreNameConverter

from dtos.store_detail_dto import StoreDetailDto

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select

import time

import pandas as pd
from openpyxl import load_workbook


class CocoblancOrderCancelProcess:
    def __init__(self):
        self.default_wait = 10
        # open_browser()
        # self.driver: webdriver.Chrome = get_chrome_driver(is_headless=False, is_secret=False)
        self.driver: webdriver.Chrome = get_chrome_driver_new(is_headless=False, is_secret=False)
        self.driver.implicitly_wait(self.default_wait)
        self.driver.maximize_window()

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
            dict_accounts[channel] = {"도메인": domain, "ID": account_id, "PW": account_pw, "URL": url}
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

    # 이지어드민 로그인 시 발생하는 팝업창을 모두 닫습니다.
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

            id_input = driver.find_element(By.XPATH, '//input[@placeholder="이메일"]')
            time.sleep(0.2)
            id_input.clear()
            time.sleep(0.2)
            id_input.send_keys(login_id)

            pw_input = driver.find_element(By.XPATH, '//input[@placeholder="비밀번호"]')
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
                EC.visibility_of_element_located((By.XPATH, '//div[contains(text(), "코코블랑")]'))
            )
            time.sleep(0.2)

            store_link = driver.find_element(By.XPATH, '//a[contains(@href, "cocoblanc")]')
            driver.execute_script("arguments[0].click();", store_link)
            time.sleep(0.2)

            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//span[contains(text(), "광고 관리")]'))
            )
            time.sleep(0.2)

        except Exception as e:
            self.log_msg.emit("지그재그 로그인 실패")
            print(e)
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
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//img[@alt="위메프 파트너 2.0"]')))
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
            time.sleep(1)

            change_next_time_button = driver.find_element(By.XPATH, '//button[contains(text(), "다음에 변경")]')
            time.sleep(0.2)
            change_next_time_button.click()
            time.sleep(1)

        except Exception as e:
            print("로그인 정보 입력 실패")

        finally:
            driver.implicitly_wait(self.default_wait)

        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//h1[./a[contains(text(), "TMON 배송상품 파트너센터")]]'))
            )
            time.sleep(0.5)

        except Exception as e:
            self.log_msg.emit(f"티몬 로그인 실패")
            print(e)
            raise Exception("티몬 로그인 실패")

    # 전체작업 시작
    def work_start(self):
        print(f"process: work_start")

        try:
            # 계정 엑셀 파일
            self.dict_accounts = self.get_dict_account()

            self.ezadmin_login()

            store_detail_dto = StoreDetailDto()

        except Exception as e:
            print(str(e))
        finally:
            self.driver.quit()
            time.sleep(0.2)


if __name__ == "__main__":
    # process = CocoblancOrderCancelProcess()
    # process.work_start()
    pass
