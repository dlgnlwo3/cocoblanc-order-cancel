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
            pass
        elif account == "쿠팡":
            self.coupang_login()
        elif account == "11번가":
            self.eleven_street_login()
        elif account == "네이버":
            pass

    def eleven_street_login(self):
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
            order_cancel_list = []
        elif account == "브리치":
            order_cancel_list = []
        elif account == "쿠팡":
            order_cancel_list = []
        elif account == "11번가":
            order_cancel_list = []
        elif account == "네이버":
            order_cancel_list = []

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

            # 주문번호 목록
            # $x('//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding"))]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_3_")]')
            order_number_list = driver.find_elements(
                By.XPATH,
                '//table[@class="gridBodyTable"]//tr[not(contains(@class, "padding"))]//div[contains(@class, "bodyTdText")][contains(@id, "AX_0_AX_3_")]',
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

            # 주문번호 목록
            # $x('//div[@id="claimCancelListGrid"]//tr[contains(@class, "dhx_web")]/td[4]')
            order_number_list = driver.find_elements(
                By.XPATH,
                '//div[@id="claimCancelListGrid"]//tr[contains(@class, "dhx_web")]/td[contains(@style, "underline")][1]',
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
                    print(account)
                elif account == "브리치":
                    print(account)
                elif account == "쿠팡":
                    print(account)
                elif account == "11번가":
                    print(account)
                elif account == "네이버":
                    print(account)

            except Exception as e:
                print(str(e))
                self.log_msg.emit(f"{account} {order_cancel_number}: 작업 실패")
                self.log_msg.emit(f"{str(e)}")
                continue

    def kakao_shopping_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

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

            # 취소 품목
            order_cancel_target = driver.find_element(
                By.XPATH,
                f'//button[contains(@onclick, "claim.popOrderDetail") and contains(@onclick, "{order_cancel_number}")]',
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
                # driver.execute_script("arguments[0].click();", kakao_order_cancel_button)
                # time.sleep(0.5)

                # 취소 승인 하시겠습니까? alert

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))

            finally:
                driver.close()
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))

        finally:
            pass

    def wemakeprice_order_cancel(self, account, order_cancel_number):
        driver = self.driver

        # 주문번호 이지어드민 검증
        try:
            self.check_order_cancel_number_from_ezadmin(account, order_cancel_number)

        except Exception as e:
            print(str(e))
            if order_cancel_number in str(e):
                raise Exception((str(e)))

        try:
            driver.get("https://wpartner.wemakeprice.com/claim/cancelMain")

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//h2[text()="취소관리"]')))
            time.sleep(0.2)

            # 300개씩 보기
            schLimitCnt_select = Select(driver.find_element(By.XPATH, '//select[@id="schLimitCnt"]'))
            schLimitCnt_select.select_by_visible_text("300개")
            time.sleep(3)

            # 취소 품목 체크박스
            # $x('//tr[./td[contains(@style, "underline") and contains(text(), "442530851")]]//img[contains(@onclick, "cancelBubble")]')
            order_cancel_target = driver.find_element(
                By.XPATH,
                f'//tr[./td[contains(@style, "underline") and contains(text(), "{order_cancel_number}")]]//img[contains(@onclick, "cancelBubble")]',
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

                order_cancel_iframe = driver.find_element(By.XPATH, '//iframe[contains(@src, "omsOrderDetail")]')
                driver.switch_to.frame(order_cancel_iframe)
                time.sleep(0.5)

                wemakeprice_order_cancel_button = driver.find_element(
                    By.XPATH, '//button[contains(text(), "취소승인(환불)")]'
                )
                # driver.execute_script("arguments[0].click();", wemakeprice_order_cancel_button)
                # time.sleep(0.5)

                # 취소 승인 하시겠습니까? alert

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))

            finally:
                driver.close()
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))

        finally:
            pass

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

            # 새 창 열림 or alert ['선택된 상품주문건이 없습니다.', '취소처리가 가능한 건이 없습니다. 클레임 상태를 확인해 주세요.']
            other_tabs = [
                tab for tab in driver.window_handles if tab != self.cs_screen_tab and tab != self.shop_screen_tab
            ]
            wemakeprice_order_cancel_tab = other_tabs[0]

            try:
                driver.switch_to.window(wemakeprice_order_cancel_tab)
                time.sleep(1)

                order_cancel_iframe = driver.find_element(By.XPATH, '//iframe[contains(@src, "omsOrderDetail")]')
                driver.switch_to.frame(order_cancel_iframe)
                time.sleep(0.5)

                wemakeprice_order_cancel_button = driver.find_element(
                    By.XPATH, '//button[contains(text(), "취소승인(환불)")]'
                )
                # driver.execute_script("arguments[0].click();", wemakeprice_order_cancel_button)
                # time.sleep(0.5)

                # 취소 승인 하시겠습니까? alert

                self.log_msg.emit(f"{account} {order_cancel_number}: 취소 완료")

            except Exception as e:
                print(str(e))

            finally:
                driver.close()
                driver.switch_to.window(self.shop_screen_tab)

        except Exception as e:
            print(str(e))

        finally:
            pass

    def check_order_cancel_number_from_ezadmin(self, account, order_cancel_number):
        driver = self.driver
        try:
            driver.switch_to.window(self.cs_screen_tab)

            if account == "네이버":
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "주문번호")]/input')
            else:
                input_order_number = driver.find_element(By.XPATH, '//td[contains(text(), "주문번호")]/input')

            input_order_number.clear()
            time.sleep(0.2)

            input_order_number.send_keys(order_cancel_number)
            time.sleep(0.2)

            search_button = driver.find_element(By.XPATH, '//div[@id="search"][text()="검색"]')
            driver.execute_script("arguments[0].click();", search_button)
            time.sleep(3)

            try:
                order_cancel_number_td = driver.find_element(By.XPATH, f'//td[@title="{order_cancel_number}"]')
                driver.execute_script("arguments[0].click();", order_cancel_number_td)
            except Exception as e:
                print(str(e))
                raise Exception(f"{account} {order_cancel_number}: 이지어드민 검색 결과가 없습니다.")
            finally:
                time.sleep(0.5)

            product_cs_state = driver.find_element(By.XPATH, '//td[@id="di_product_cs"]')
            product_cs_state = product_cs_state.get_attribute("textContent")
            print(f"{account}, {order_cancel_number}, {product_cs_state}")

            if not "배송전 주문취소" in product_cs_state:
                raise Exception(f"{account} {order_cancel_number}: 배송전 주문취소 상태가 아닙니다.")

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

        finally:
            driver.quit()
            time.sleep(0.2)


if __name__ == "__main__":
    # process = CocoblancOrderCancelProcess()
    # process.work_start()
    pass
