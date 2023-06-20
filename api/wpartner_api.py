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


class WPartnerAPI:
    def __init__(self, apiKey):
        self.apiKey = apiKey
        self.get_headers()

    def get_headers(self):
        self.headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "apiKey": f"{self.apiKey}",
        }

    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def getOrderList_from_date(self, fromDate, toDate):
        auth_url = f"https://wapi-stg.wemakeprice.com/order/out/getOrderList"
        params = {
            "fromDate": fromDate,
            "toDate": toDate,
            "type": "NEW",
            "searchDateType": "NEW",
        }

        result = requests.post(auth_url, headers=self.headers, params=params)
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

    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def get_notice_from_wemakeprice(self):
        auth_url = f"https://w-api.wemakeprice.com/notice/out/getNoticeList"
        params = {"basicDate": "2023-06-08"}

        result = requests.get(auth_url, headers=self.headers, params=params)
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
    apiKey = "4a8197bc0dff7df87a2419b984a729842fed49288f9b682e406e9cab75665146a54ad7823dce4cf1dd6c16a943be1ecde352963c97accd00441af381f19a660d"

    APIBot = WPartnerAPI(apiKey)

    fromDate = "2023-06-07 00:00:00"
    toDate = "2023-06-07 23:59:59"

    data = asyncio.run(APIBot.get_notice_from_wemakeprice())

    print(type(data))

    print(data)

    print(len(data))

    clipboard.copy(str(data))
