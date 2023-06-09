if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

from enum import Enum


class StoreNameEnum(Enum):
    Cafe24 = "카페24"
    ElevenStreet = "11번가"
    KakaoTalkStore = "카카오톡스토어"
    GMarket = "G마켓"
    Brandi = "브랜디"
    Brich = "브리치"
    WeMakePrice = "위메이크프라이스 2.0"
    Coupang = "쿠팡(자동)"
    TicketMonster = "티몬(자동)"
    ZigZag = "지그재그"


if __name__ == "__main__":
    print(StoreNameEnum.Cafe24.value)
