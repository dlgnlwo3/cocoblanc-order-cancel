if 1 == 1:
    import sys
    import warnings
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
    warnings.simplefilter("ignore", UserWarning)
    sys.coinit_flags = 2

from enums.store_name_enum import StoreNameEnum


class StoreNameConverter:
    def __init__(self):
        print(f"상점 이름 변환")

    # 엑셀에 적혀있는 상점의 이름을 이지어드민에 있는 이름으로 변환해야 합니다.
    def convert_store_name(self, store_name: str):
        print(f"{store_name}")

        if store_name == "11번가":
            store_name = StoreNameEnum.ElevenStreet.value
        elif store_name == "지마켓":
            store_name = StoreNameEnum.GMarket.value
        elif store_name == "브랜디":
            store_name = StoreNameEnum.Brandi.value
        elif store_name == "브리치":
            store_name = StoreNameEnum.Brich.value
        elif store_name == "위메프":
            store_name = StoreNameEnum.WeMakePrice.value
        elif store_name == "카카오":
            store_name = StoreNameEnum.KakaoTalkStore.value
        elif store_name == "자사몰(까페24기준)" or store_name == "카페24" or store_name == "까페24":
            store_name = StoreNameEnum.Cafe24.value
        elif store_name == "쿠팡":
            store_name = StoreNameEnum.Coupang.value
        elif store_name == "티몬":
            store_name = StoreNameEnum.TicketMonster.value

        return store_name
