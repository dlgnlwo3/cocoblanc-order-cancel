if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

from enum import Enum


class StoreNameEnum(Enum):
    ezadmin = "이지어드민"
    KakaoTalkStore = "카카오톡스토어"
    WeMakePrice = "위메프"
    TicketMonster = "티몬"
    ZigZag = "지그재그"
    Brich = "브리치"
    Coupang = "쿠팡"
    ElevenStreet = "11번가"
    Naver = "네이버"


if __name__ == "__main__":
    print(StoreNameEnum.Naver.value)
