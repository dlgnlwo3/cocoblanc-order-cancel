if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

from enum import Enum


class ShopNameEnum(Enum):
    KakaoTalkStore = "카카오톡스토어"
    WeMakePrice = "위메프"
    TicketMonster = "티몬"
    Zigzag = "지그재그"
    Bflow = "브리치"
    Coupang = "쿠팡"
    Elevenst = "11번가"
    Naver = "네이버"

    shop_list = [KakaoTalkStore, WeMakePrice, TicketMonster, Zigzag, Bflow, Coupang, Elevenst, Naver]


if __name__ == "__main__":
    print(ShopNameEnum.shop_list.value)
