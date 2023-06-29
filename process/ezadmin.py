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
