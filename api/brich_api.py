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
import hmac, hashlib
import urllib.parse
import urllib.request
import ssl

from tenacity import retry, wait_fixed, stop_after_attempt


class BrichAPI:
    def __init__(self, token):
        self.token: str = token
        self.get_headers()

    def get_headers(self):
        self.headers = {"accept": "application/json", "Authorization": f"Bearer {self.token}"}

    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def get_category(self):
        auth_url = f"https://openapi.brich.co.kr/v1.1/category"

        result = requests.get(auth_url, headers=self.headers)
        result_text = result.text
        result_json = json.loads(result_text)
        print(result_json)
        print(f"status_code: {result.status_code}")

        # 200
        if result.status_code == HTTPStatus.OK:
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
    token = "e811f0fb5d19d8bcac39ac612d2817ddd79593da2749a9517f72f58ad57e4864"

    APIBot = BrichAPI(token)

    data = asyncio.run(APIBot.get_category())

    print(type(data))

    print(data)

    print(len(data))

    clipboard.copy(str(data))
