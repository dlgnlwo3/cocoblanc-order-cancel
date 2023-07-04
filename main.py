if 1 == 1:
    import sys
    import warnings
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
    warnings.simplefilter("ignore", UserWarning)
    sys.coinit_flags = 2

from PySide6.QtGui import *
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from common.chrome import *
from dtos.gui_dto import *
from common.utils import global_log_append
from tabs.cocoblanc_order_cancel_tab import CocoblancOrderCancelTab
from PySide6.QtNetwork import QNetworkAccessManager, QNetworkRequest

from configs.program_config import ProgramConfig as Config


# 오류 발생 시 프로그램 강제종료 방지
def my_exception_hook(exctype, value, traceback):
    print(exctype, value, traceback)
    global_log_append(str(value))
    sys._excepthook(exctype, value, traceback)


sys.excepthook = my_exception_hook

# pyinstaller -n "코코블랑 취소요청 승인 v0.0.6 (이지어드민 검증 과정 수정)" -w --onefile --clean "main.py" --icon "assets\ezadmin.ico"


class MainUI(QWidget):
    # 초기화
    def __init__(self):
        self.config = Config()

        print(f"exe_path: {self.config.exe_path}")
        print(f"log_folder_path: {self.config.log_folder_path}")

        # UI
        super().__init__()
        self.initUI()

    # 가운데 정렬
    def center(self):
        qr = self.frameGeometry()
        cp = QGuiApplication.primaryScreen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    # 프로그램 닫기 클릭 시
    def closeEvent(self, event):
        quit_msg = "프로그램을 종료하시겠습니까?"
        reply = QMessageBox.question(self, "프로그램 종료", quit_msg, QMessageBox.Yes, QMessageBox.No)

        if reply == QMessageBox.Yes:
            print(f"프로그램을 종료합니다.")
            event.accept()
        else:
            print(f"종료 취소")
            event.ignore()

    # 아이콘 설정
    def set_window_icon_from_response(self, http_response):
        pixmap = QPixmap()
        pixmap.loadFromData(http_response.readAll())
        icon = QIcon(pixmap)
        self.setWindowIcon(icon)

    # 메인 UI
    def initUI(self):
        # 이미지 주소를 복사해야 함
        ICON_IMAGE_URL = "https://i.imgur.com/bV1HDvH.png"
        self.icon = QNetworkAccessManager()
        self.icon.finished.connect(self.set_window_icon_from_response)
        self.icon.get(QNetworkRequest(QUrl(ICON_IMAGE_URL)))

        # 탭 초기화
        self.cocoblanc_order_cancel_tab = CocoblancOrderCancelTab()

        # 탭 추가
        tabs = QTabWidget()
        tabs.addTab(self.cocoblanc_order_cancel_tab, "취소요청승인")

        vbox = QVBoxLayout()

        vbox.addWidget(tabs)
        self.setLayout(vbox)

        # 앱 기본 설정
        self.setWindowTitle(f"코코블랑 취소요청 승인 v0.0.6")
        self.resize(600, 700)
        self.center()
        self.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MainUI()
    sys.exit(app.exec_())
