if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


class ProductDto:
    def __init__(self):
        self.__order_number = ""
        self.__product_name = ""
        self.__product_option = ""
        self.__product_qty = ""
        self.__product_recv_name = ""
        self.__product_recv_tel = ""

    @property
    def order_number(self):  # getter
        return self.__order_number

    @order_number.setter
    def order_number(self, value):  # setter
        self.__order_number = value

    @property
    def product_name(self):  # getter
        return self.__product_name

    @product_name.setter
    def product_name(self, value):  # setter
        self.__product_name = value

    @property
    def product_option(self):  # getter
        return self.__product_option

    @product_option.setter
    def product_option(self, value):  # setter
        self.__product_option = value

    @property
    def product_qty(self):  # getter
        return self.__product_qty

    @product_qty.setter
    def product_qty(self, value):  # setter
        self.__product_qty = value

    @property
    def product_recv_name(self):  # getter
        return self.__product_recv_name

    @product_recv_name.setter
    def product_recv_name(self, value):  # setter
        self.__product_recv_name = value

    @property
    def product_recv_tel(self):  # getter
        return self.__product_recv_tel

    @product_recv_tel.setter
    def product_recv_tel(self, value):  # setter
        self.__product_recv_tel = value

    def get_dict(self) -> dict:
        return {
            "주문번호": self.order_number,
            "상품명": self.product_name,
            "상품옵션": self.product_option,
            "수량": self.product_qty,
            "수령자명": self.product_recv_name,
            "수령자연락처": self.product_recv_tel,
        }

    def to_print(self):
        print("주문번호: ", self.order_number)
        print("상품명: ", self.product_name)
        print("상품옵션: ", self.product_option)
        print("수량: ", self.product_qty)
        print("수령자명: ", self.product_recv_name)
        print("수령자연락처: ", self.product_recv_tel)
