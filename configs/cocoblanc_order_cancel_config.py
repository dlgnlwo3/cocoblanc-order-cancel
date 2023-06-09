if 1 == 1:
    import sys
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import os
import json
from configs.program_config import ProgramConfig


class CocoblancOrderCancelData:
    def __init__(self):
        self.account_file = ""
        self.stats_file = ""

    def __member__(self):
        members = []
        for member in dir(self):
            if not member.startswith("__"):
                members.append(member)
        return members


class CocoblancOrderCancelConfig(ProgramConfig):
    config_name = "cocoblanc_order_cancel"

    def __init__(self):
        ProgramConfig.__init__(self)
        self.save_path = os.path.join(self.program_path, f"{self.program_id}_{self.config_name}.txt")
        self.init_data()

    # # 저장정보
    def init_data(self):
        # 저장데이터 없는 경우 초기화
        member_datas = CocoblancOrderCancelData().__member__()
        saved_data = self.get_data()
        new_write_data = {}

        # 저장된 데이터가 하나도 없는 경우 ""으로 초기화
        for member in member_datas:
            if not member in saved_data.keys():
                # 초기화
                new_write_data.update({member: ""})
            else:
                exist_value = saved_data[member]
                new_write_data.update({member: exist_value})

        self.write_data(new_write_data)

    def data_to_dict(self, data: CocoblancOrderCancelData) -> dict:
        member_datas = data.__member__()
        dict_data = {}

        # 저장된 데이터가 하나도 없는 경우 ""으로 초기화
        for member in member_datas:
            dict_data.update({member: data.__getattribute__(member)})

        return dict_data

    def dict_to_data(self, data: dict) -> CocoblancOrderCancelData:
        newData = CocoblancOrderCancelData()
        for key in data.keys():
            newData.__setattr__(key, data[key])
        return newData

    def write_data(self, save_data: dict):
        # 지우고 다시 저장
        if os.path.isfile(self.save_path):
            os.remove(self.save_path)

        with open(self.save_path, "w", encoding=self.encoding) as f:
            f.write(json.dumps((save_data)))
            f.close()

        return save_data

    def get_data(self) -> dict:
        saved_data = {}
        if os.path.isfile(self.save_path):
            with open(self.save_path, "r", encoding=self.encoding) as f:
                saved_data = json.loads(f.read())

        member_datas = CocoblancOrderCancelData().__member__()
        for member in member_datas:
            if not member in saved_data.keys():
                # 초기화
                saved_data.update({member: ""})
            else:
                exist_value = saved_data[member]
                saved_data.update({member: exist_value})
        return saved_data


if __name__ == "__main__":
    addData = CocoblancOrderCancelData()
    CocoblancOrderCancelConfig = CocoblancOrderCancelConfig()
    dict_data = CocoblancOrderCancelConfig.data_to_dict(addData)
    print(dict_data)
