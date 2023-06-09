class StoreDetailDto:
    def __init__(self):
        self.__store_name = ""

        self.__tot_products = ""  # 주문수량

        self.__tot_amount = ""  # 주문금액

        self.__org_price = ""  # 상품원가

        self.__cancel_total_data_product_sum = ""  # 취소수량

        self.__cancel_total_data_order_sum_amount = ""  # 취소금액

        self.__refund_total_data_product_sum = ""  # 반품수량

        self.__refund_total_data_order_sum_amount = ""  # 반품금액

        self.__delivery_result = ""  # 배송건수

        self.__zigzag_cost = ""  # 지그재그

        self.__mypick_cost = ""  # 마이픽쿠폰

        self.__coupon_cost = ""  # 쿠폰비

        self.__tot_supply_price = ""  # 정산예정금액

        self.__cancel_amount_price = ""  # 판매취소금액

        self.__cancel_price = ""  # 정산취소금액

        # 카페24용 당일배송 관련
        self.__today_delivery_result = ""

    @property
    def store_name(self):  # getter
        return self.__store_name

    @store_name.setter
    def store_name(self, value: str):  # setter
        self.__store_name = value

    @property
    def tot_products(self):  # getter
        return self.__tot_products

    @tot_products.setter
    def tot_products(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__tot_products = int_value

    @property
    def tot_amount(self):  # getter
        return self.__tot_amount

    @tot_amount.setter
    def tot_amount(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__tot_amount = int_value

    @property
    def org_price(self):  # getter
        return self.__org_price

    @org_price.setter
    def org_price(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__org_price = int_value

    @property
    def cancel_total_data_product_sum(self):  # getter
        return self.__cancel_total_data_product_sum

    @cancel_total_data_product_sum.setter
    def cancel_total_data_product_sum(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__cancel_total_data_product_sum = int_value

    @property
    def cancel_total_data_order_sum_amount(self):  # getter
        return self.__cancel_total_data_order_sum_amount

    @cancel_total_data_order_sum_amount.setter
    def cancel_total_data_order_sum_amount(self, value: str):  # setter
        int_value = 0
        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__cancel_total_data_order_sum_amount = int_value

    @property
    def refund_total_data_product_sum(self):  # getter
        return self.__refund_total_data_product_sum

    @refund_total_data_product_sum.setter
    def refund_total_data_product_sum(self, value: str):  # setter
        int_value = 0
        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__refund_total_data_product_sum = int_value

    @property
    def refund_total_data_order_sum_amount(self):  # getter
        return self.__refund_total_data_order_sum_amount

    @refund_total_data_order_sum_amount.setter
    def refund_total_data_order_sum_amount(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__refund_total_data_order_sum_amount = int_value

    @property
    def delivery_result(self):  # getter
        return self.__delivery_result

    @delivery_result.setter
    def delivery_result(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__delivery_result = int_value

    @property
    def today_delivery_result(self):  # getter
        return self.__today_delivery_result

    @today_delivery_result.setter
    def today_delivery_result(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__today_delivery_result = int_value

    @property
    def zigzag_cost(self):  # getter
        return self.__zigzag_cost

    @zigzag_cost.setter
    def zigzag_cost(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__zigzag_cost = int_value

    @property
    def mypick_cost(self):  # getter
        return self.__mypick_cost

    @mypick_cost.setter
    def mypick_cost(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__mypick_cost = int_value

    @property
    def coupon_cost(self):  # getter
        return self.__coupon_cost

    @coupon_cost.setter
    def coupon_cost(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__coupon_cost = int_value

    @property
    def tot_supply_price(self):  # getter
        return self.__tot_supply_price

    @tot_supply_price.setter
    def tot_supply_price(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__tot_supply_price = int_value

    @property
    def cancel_amount_price(self):  # getter
        return self.__cancel_amount_price

    @cancel_amount_price.setter
    def cancel_amount_price(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__cancel_amount_price = int_value

    @property
    def cancel_price(self):  # getter
        return self.__cancel_price

    @cancel_price.setter
    def cancel_price(self, value: str):  # setter
        int_value = 0

        if value == 0:
            pass

        elif value != "":
            int_value = value.replace(",", "")

        try:
            int_value = int(int_value)
        except Exception as e:
            int_value = 0

        self.__cancel_price = int_value

    def to_print(self):
        print("상점이름", self.store_name)
