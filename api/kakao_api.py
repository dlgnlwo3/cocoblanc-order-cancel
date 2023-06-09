#!/usr/bin/env python
if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

from http import HTTPStatus
import bcrypt
import pybase64
import requests
import json
import time
import asyncio
import clipboard

from tenacity import retry, wait_fixed, stop_after_attempt

import datetime


class KakaoAPI:
    def __init__(self, NATIVE_APP_KEY, REST_API_KEY, JAVASCRIPT_KEY, ADMIN_KEY, CHANNEL_ID, SELLER_API_KEY):
        self.NATIVE_APP_KEY = NATIVE_APP_KEY
        self.REST_API_KEY = REST_API_KEY
        self.JAVASCRIPT_KEY = JAVASCRIPT_KEY
        self.ADMIN_KEY = ADMIN_KEY
        self.CHANNEL_ID = CHANNEL_ID
        self.SELLER_API_KEY = SELLER_API_KEY
        self.get_headers()

    def get_headers(self):
        self.headers = {
            "Content-Type": "application/json;charset=UTF-8",
            # "Host": "kapi.kakao.com",
            "Authorization": f"KakaoAK {self.ADMIN_KEY}",
            "Target-Authorization": f"KakaoAK {self.SELLER_API_KEY}",
            "channel-ids": self.CHANNEL_ID,
        }

    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def connect_kakao_seller_and_developer(self):
        auth_url = "https://kapi.kakao.com/v1/store/register"
        result = requests.post(auth_url, headers=self.headers)
        result_text = result.text.encode("utf-8")
        result_json = json.loads(result_text)
        print(result_json)
        print(f"status_code: {result.status_code}")

        # 200
        if result.status_code == HTTPStatus.OK:
            result_json = result_json["orderNumberList"]
            print("성공")

        # 400
        elif result.status_code == HTTPStatus.BAD_REQUEST:
            print("입력값이 유효하지 않음")

        # 404
        elif result.status_code == HTTPStatus.NOT_FOUND:
            print("Request-URI에 일치하는 건을 발견하지 못함")

        # 405
        elif result.status_code == HTTPStatus.METHOD_NOT_ALLOWED:
            print("허가되지 않은 메소드 사용")

        # 500
        elif result.status_code == HTTPStatus.INTERNAL_SERVER_ERROR:
            print("서버 내부의 에러")

        # 503
        elif result.status_code == HTTPStatus.SERVICE_UNAVAILABLE:
            print("서버 과부하로 인한 사용 불가")

        # 그 외의 경우
        else:
            print("알 수 없는 오류")

        return result_json

    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def get_seller_address(self):
        auth_url = "https://kapi.kakao.com/v1/shopping/bizseller/seller-addresses/search"
        result = requests.get(auth_url, headers=self.headers)
        result_text = result.text.encode("utf-8")
        result_json = json.loads(result_text)
        print(result_json)
        print(f"status_code: {result.status_code}")

        # 200
        if result.status_code == HTTPStatus.OK:
            result_json = result_json["orderNumberList"]
            print("성공")

        # 400
        elif result.status_code == HTTPStatus.BAD_REQUEST:
            print("입력값이 유효하지 않음")

        # 404
        elif result.status_code == HTTPStatus.NOT_FOUND:
            print("Request-URI에 일치하는 건을 발견하지 못함")

        # 405
        elif result.status_code == HTTPStatus.METHOD_NOT_ALLOWED:
            print("허가되지 않은 메소드 사용")

        # 500
        elif result.status_code == HTTPStatus.INTERNAL_SERVER_ERROR:
            print("서버 내부의 에러")

        # 503
        elif result.status_code == HTTPStatus.SERVICE_UNAVAILABLE:
            print("서버 과부하로 인한 사용 불가")

        # 그 외의 경우
        else:
            print("알 수 없는 오류")

        return result_json


if __name__ == "__main__":
    # kakao developers
    NATIVE_APP_KEY = "c2d706f75195cb1cc0cad735243e8b83"
    REST_API_KEY = "9f3beef01ca2c28ef03aed8cba0bc44d"
    JAVASCRIPT_KEY = "c758f40fc2d0d1ab0ab0b895094ceb0b"
    ADMIN_KEY = "fc58e4eb677d1b74cff6983e3a6f10c0"
    CHANNEL_ID = "101"

    # kakao_sellers
    SELLER_API_KEY = "171e14047402a0cbb462a9b6700b19cf"

    APIBot = KakaoAPI(NATIVE_APP_KEY, REST_API_KEY, JAVASCRIPT_KEY, ADMIN_KEY, CHANNEL_ID, SELLER_API_KEY)

    data = ""

    # 날짜를 이용해서 주문번호 목록을 가져옵니다.
    # data = asyncio.run(APIBot.get_searchOrder_list_from_date(start_date, end_date))

    # 주문번호를 이용해서 주문상세정보를 가져옵니다.
    data = asyncio.run(APIBot.connect_kakao_seller_and_developer())

    print(type(data))

    print(data)

    print(len(data))

    clipboard.copy(str(data))
