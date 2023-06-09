if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


class GUIDto:
    def __init__(self):
        self.__account_file = ""
        self.__stats_file = ""
        self.__sheet_name = ""

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

    def to_print(self):
        print("account_file: ", self.account_file)
        print("stats_file: ", self.stats_file)
        print("sheet_name: ", self.sheet_name)
