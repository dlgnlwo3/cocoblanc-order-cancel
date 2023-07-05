if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


class GUIDto:
    def __init__(self):
        self.__account_file = ""
        self.__stats_file = ""
        self.__sheet_name = ""

        self.__selected_shop_list = []

    @property
    def account_file(self):  # getter
        return self.__account_file

    @account_file.setter
    def account_file(self, value):  # setter
        self.__account_file = value

    @property
    def stats_file(self):  # getter
        return self.__stats_file

    @stats_file.setter
    def stats_file(self, value):  # setter
        self.__stats_file = value

    @property
    def sheet_name(self):  # getter
        return self.__sheet_name

    @sheet_name.setter
    def sheet_name(self, value):  # setter
        self.__sheet_name = value

    @property
    def selected_shop_list(self):  # getter
        return self.__selected_shop_list

    @selected_shop_list.setter
    def selected_shop_list(self, value):  # setter
        self.__selected_shop_list = value

    def to_print(self):
        print("account_file: ", self.account_file)
        print("stats_file: ", self.stats_file)
        print("sheet_name: ", self.sheet_name)
        print("selected_shop_list: ", self.selected_shop_list)
