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

from datetime import datetime

from tenacity import retry, wait_fixed, stop_after_attempt


class CoupangAPI:
    def __init__(self, vendorId, Access_Key, Secret_Key):
        os.environ["TZ"] = "GMT+0"
        self.vendorId: str = vendorId
        self.Access_Key: str = Access_Key
        self.Secret_Key: str = Secret_Key
        self.now = datetime.now().strftime("%y%m%dT%H%M%SZ")

    def get_headers(self, path, query):
        message = self.now + "GET" + path + query
        signature = hmac.new(self.Secret_Key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).hexdigest()
        authorization = (
            "CEA algorithm=HmacSHA256, access-key="
            + self.Access_Key
            + ", signed-date="
            + self.now
            + ", signature="
            + signature
        )

        self.headers = {
            "Content-type": "application/json;charset=UTF-8",
            "Authorization": authorization,
        }

    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def get_returnRequests_from_date(self, createdAtFrom, createdAtTo):
        url = f"https://api-gateway.coupang.com/v2/providers/openapi/apis/api/v4/vendors/{self.vendorId}/returnRequests"
        query = urllib.parse.urlencode({"createdAtFrom": createdAtFrom, "createdAtTo": createdAtTo, "status": "UC"})
        full_url = url + "?" + query

        headers = self.get_headers(path=url, query=query)

        result = requests.get(full_url, headers=self.headers)
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
    vendorId = "A00121999"
    Access_Key = "150067ce-399b-4aad-90c2-c0537ab86736"
    Secret_Key = "db2bc4d642d3a53305345446a61ad19d64c732f7"

    APIBot = CoupangAPI(vendorId, Access_Key, Secret_Key)

    createdAtFrom = "2023-06-06"
    createdAtTo = "2023-06-07"

    # 주문번호를 이용해서 주문상세정보를 가져옵니다.
    data = asyncio.run(APIBot.get_returnRequests_from_date(createdAtFrom, createdAtTo))

    print(type(data))

    print(data)

    print(len(data))

    clipboard.copy(str(data))
