import pyperclip
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
import time
from selenium import webdriver

# element로 마우스 이동  : ActionChains(driver).move_to_element(ref)
# element 마우스 클릭	: ActionChains(driver).click(ref)
# element 키보드 입력	: ActionChains(driver).send_keys_to_element(ref, keys)
# 키보드 입력	        :  ActionChains(driver).send_keys(keys)
# links = driver.find_elements_by_xpath('//li/a/@href')
# driver.find_element_by_css_selector(f'ul.send_name > li[data-value="{window.mail_sender}"]').click()
# driver.find_element_by_css_selector("button[data-name='align-drop-down-with-justify']").click()
# find_element_by_id
# driver.find_elements_by_css_selector('a[href*="https://cr.shopping.naver.com/adcr.nhn"][class*="basicList_link"]')

# xpath 규칙 참조
# / : 절대경로를 나타냄
# // : 문서내에서 검색
# //@href : href 속성이 있는 모든 태그 선택
# //a[@href='http://google.com'] : a 태그의 href 속성에 http://google.com 속성값을 가진 모든 태그 선택
# (//a)[3] : 문서의 세 번째 링크 선택
# (//table)[last()] : 문서의 마지막 테이블 선택
# (//a)[position() < 3] : 문서의 처음 두 링크 선택
# //table/tr/* 모든 테이블에서 모든 자식 tr 태그 선택
# //div[@*] 속성이 하나라도 있는 div 태그 선택

#  아래 방법으로 크롬에서 xpath 가져오기 연습 가능
# function getElementByXpath(path) {
#   return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
# }

# console.log( getElementByXpath("//html[1]/body[1]/div[1]") );


def switch_to_default_frame(driver: webdriver.Chrome):
    driver.switch_to.default_content()


def switch_to_iframe_by_css_selector(driver, css_selector):
    frame = driver.find_element_by_css_selector(css_selector)
    driver.switch_to.frame(frame)


def switch_to_iframe_by_xpath(driver, xpath):
    frame = driver.find_element_by_xpath(xpath)
    driver.switch_to.frame(frame)


def triple_click(driver):
    for i in range(3):
        actions = ActionChains(driver)
        actions.click()
        actions.perform()


def alert_ok_try(driver):
    try:
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass


def alert_ok(driver):
    alert = driver.switch_to.alert
    alert.accept()


def type_into_by_element(el, text):
    el.click()
    el.clear()
    pyperclip.copy(text)
    el.send_keys(Keys.CONTROL, "v")


# 로그인 정보 입력 함수


def clipboard_input(driver, target_xpath, text):
    temp_user_input = pyperclip.paste()
    pyperclip.copy(text)
    driver.find_element_by_xpath(target_xpath).click()
    ActionChains(driver).key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
    pyperclip.copy(temp_user_input)
    time.sleep(1)


def type_into_by_xpath(driver, target_xpath, text):
    el = driver.find_element_by_xpath(target_xpath)
    el.click()
    el.clear()
    pyperclip.copy(text)
    el.send_keys(Keys.CONTROL, "v")


def type_into_by_id(driver, target_id, text):
    el = driver.find_element_by_id(target_id)
    el.click()
    el.clear()
    pyperclip.copy(text)
    el.send_keys(Keys.CONTROL, "v")


def type_into_by_class_name(driver, class_name, text):
    el = driver.find_element_by_class_name(class_name)
    el.click()
    el.clear()
    pyperclip.copy(text)
    el.send_keys(Keys.CONTROL, "v")


def type_into_by_class_name2(driver, class_name, text):
    el = driver.find_element_by_class_name(class_name)
    el.click()
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.DELETE)
    pyperclip.copy(text)
    el.send_keys(Keys.CONTROL, "v")


def type_into_by_tag_name(driver, tag_name, text):
    el = driver.find_element(By.TAG_NAME, tag_name)
    el.click()
    el.clear()
    pyperclip.copy(text)
    el.send_keys(Keys.CONTROL, "v")


def type_into_by_tag_name_first_children(driver, tag_name, text):
    el = driver.find_elements(By.TAG_NAME, tag_name)[0]
    el.click()
    el.clear()
    pyperclip.copy(text)
    el.click()
    el.send_keys(Keys.CONTROL, "v")


def paste_image_by_tag_name(driver, tag_name):
    el = driver.find_elements(By.TAG_NAME, tag_name)[0]
    el.click()
    el.send_keys(Keys.CONTROL, "v")

    el_caption = driver.find_elements(By.TAG_NAME, "figcaption")[0]
    el_caption.send_keys(Keys.ENTER)


def click_if_xpath_exist(driver, xpath, seconds=30):
    driver.implicitly_wait(seconds)
    try:
        el = driver.find_element_by_xpath(xpath)
        el.click()
    except:
        pass


def click_down_and_up(driver, target_xpath):
    element_to_click = driver.find_element_by_xpath(target_xpath)
    actions = ActionChains(driver)
    actions.move_to_element(element_to_click).click_and_hold(element_to_click).perform()


def double_click(driver, target_xpath):
    element_to_click = driver.find_element_by_xpath(target_xpath)
    actionChains = ActionChains(driver)
    actionChains.double_click(element_to_click).perform()


def move_and_click(driver, target_path):
    element = driver.find_element_by_xpath(target_path)
    action = ActionChains(driver)
    action.move_to_element(element).perform()
    element.click()


def send_keys_to_driver(driver, key):
    actions = ActionChains(driver)
    actions.send_keys(key)
    actions.perform()


def send_multikeys_to_driver(driver, key1, key2):
    ActionChains(driver).key_down(key1).send_keys(key2).key_up(key1).perform()


def click_text(driver, text):
    driver.find_element_by_xpath(f"//*[text()='{text}']")


def click_tag_by_text(driver, tag, text):
    wait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//{tag}[text()='{text}']"))).click()


def click_tag_by_text_with_wait(driver, tag, text, wait_second):
    wait(driver, wait_second).until(EC.element_to_be_clickable((By.XPATH, f"//{tag}[text()='{text}']"))).click()


def type_into_tag_by_text(driver, tag, text, content):
    wait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//{tag}[text()='{text}']"))).send_keys(content)


def click_tag_by_contain_text(driver, tag, text):
    wait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//{tag}[contains(text(), "{text}")]'))).click()


# ex) '//div[@class="column_category"]//button[contains(text(), "우니유로")]'


def click_tag_by_contain_text_with_parent_node(driver, parent_node, tag, text):
    wait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f'//{parent_node}//{tag}[contains(text(), "{text}")]'))
    ).click()


def mouse_over_to_element_by_class_name(driver, class_name):
    a = ActionChains(driver)
    m = driver.find_element_by_class_name(class_name)
    a.move_to_element(m).perform()


def mouse_over_to_element_by_element(driver, element):
    a = ActionChains(driver)
    a.move_to_element(element).perform()


def scroll_down(driver, scroll_down_count):
    for i in range(scroll_down_count + 1):
        driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.PAGE_DOWN)
        time.sleep(0.1)


def scroll_down_old(driver, scroll_down_count):
    for i in range(scroll_down_count):
        driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.PAGE_DOWN)
        time.sleep(1)

    last_page_height = driver.execute_script("return document.documentElement.scrollHeight")
    for i in range(scroll_down_count):
        driver.execute_script("window.scrollTo(0, document.documentElement.scrollHeight);")
        time.sleep(2)
        new_page_height = driver.execute_script("return document.documentElement.scrollHeight")
        if new_page_height == last_page_height:
            break
        last_page_height = new_page_height


def page_down(driver, page_down_count):
    for i in range(page_down_count):
        send_keys_to_driver(driver, Keys.PAGE_DOWN)
        time.sleep(0.5)


def key_down(driver, down_count):
    for i in range(down_count):
        send_keys_to_driver(driver, Keys.DOWN)
        time.sleep(0.5)


def close_new_tabs(driver):
    tabs = driver.window_handles
    while len(tabs) != 1:
        driver.switch_to.window(tabs[1])
        driver.close()
        tabs = driver.window_handles
    driver.switch_to.window(tabs[0])


def scroll_down_to_end(driver):
    prev_height = driver.execute_script("return document.body.scrollHeight")

    # 웹페이지 맨 아래까지 무한 스크롤
    while True:
        # 스크롤을 화면 가장 아래로 내린다
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")

        # 현재 문서 높이를 가져와서 저장
        curr_height = driver.execute_script("return document.body.scrollHeight")

        if curr_height == prev_height:
            break
        else:
            prev_height = driver.execute_script("return document.body.scrollHeight")


def scroll_down_to_end_with_wait_second(driver, wait_second):
    prev_height = driver.execute_script("return document.body.scrollHeight")

    # 웹페이지 맨 아래까지 무한 스크롤
    while True:
        # 스크롤을 화면 가장 아래로 내린다
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")

        # 페이지 로딩 대기
        time.sleep(wait_second)

        # 현재 문서 높이를 가져와서 저장
        curr_height = driver.execute_script("return document.body.scrollHeight")

        if curr_height == prev_height:
            break
        else:
            prev_height = driver.execute_script("return document.body.scrollHeight")
