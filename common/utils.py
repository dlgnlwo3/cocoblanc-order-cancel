import random
import time
from datetime import datetime
import os
from mimetypes import MimeTypes
from configs.program_config import ProgramConfig


def get_mime_type(file_path):
    mime = MimeTypes()
    mime_type, encoding = mime.guess_type(file_path)
    return mime_type


# 전역 로그
def global_log_append(text):
    text = str(text)

    today = str(datetime.now())[0:10]
    now = str(datetime.now())[0:-7]

    log_path = ProgramConfig().log_folder_path

    today_log = os.path.join(log_path, f"{today}.txt")
    if os.path.isfile(today_log) == False:
        f = open(today_log, "w", encoding="UTF8")
    else:
        f = open(today_log, "a", encoding="UTF8")
    f.write(f"[{now}] {text}\n")
    f.close()


def random_delay(start, end):
    x = random.randint(start, end)
    time.sleep(x)
