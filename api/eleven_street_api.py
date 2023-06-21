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

import xmltodict


class ElevenStreetAPI:
    def __init__(self, openapikey):
        self.openapikey = openapikey
        self.get_headers()

    def get_headers(self):
        self.headers = {"openapikey": f"{self.openapikey}"}

    # 주문 취소 목록
    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def get_cancelorders_from_date(self, startTime, endTime):
        auth_url = f"http://api.11st.co.kr/rest/claimservice/cancelorders/{startTime}/{endTime}"
        result = requests.get(auth_url, headers=self.headers)
        result_text = result.text
        result_json = xmltodict.parse(result_text)
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

    # 주문취소 승인처리
    @retry(
        wait=wait_fixed(3),  # 3초 대기
        stop=stop_after_attempt(2),  # 2번 재시도
    )
    async def cancelreqconf_from_ordInfo(self, ordPrdCnSeq, ordNo, ordPrdSeq):
        auth_url = f"http://api.11st.co.kr/rest/claimservice/cancelreqconf/{ordPrdCnSeq}/{ordNo}/{ordPrdSeq}"
        result = requests.get(auth_url, headers=self.headers)
        result_text = result.text
        result_json = xmltodict.parse(result_text)
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
    openapikey = "a69d9ede007208255aaa2f8c602d276c"

    APIBot = ElevenStreetAPI(openapikey)

    startTime = "202306210000"
    endTime = "202306212359"

    ordPrdCnSeq = ""  # 클레임번호
    ordNo = ""  # 주문번호
    ordPrdSeq = ""  # 주문순번

    # 주문번호를 이용해서 주문취소목록을 가져옵니다. (클레임 목록)
    data = asyncio.run(APIBot.get_cancelorders_from_date(startTime, endTime))

    print(type(data))

    print(data)

    print(len(data))

    clipboard.copy(str(data))
