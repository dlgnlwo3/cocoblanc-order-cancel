if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import os
import pandas as pd


class AccountFile:
    def __init__(self, filepath):
        self.filepath = filepath
        self.initData()

    def account_data_type(self):
        return {"채널명": str, "도메인": str, "ID": str, "PW": str, "URL": str}

    def initData(self):
        self.filepath = os.path.normpath(self.filepath)
        self.filename = os.path.basename(self.filepath)

        account_columns = self.account_data_type()
        try:
            self.df_account = pd.read_excel(self.filepath, converters=account_columns, keep_default_na="")
            self.df_account = self.df_account.loc[:, list(account_columns.keys())]
        except Exception as e:
            print(e)

    def output(self):
        print(self.df_account)


if __name__ == "__main__":
    synonym_file = AccountFile(r"D:\Consolework\ezadmin-crawler\excel\사이트 계정정보.xlsx")
    synonym_file.output()
