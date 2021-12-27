from PyQt5.QtGui import QPixmap, QPalette
from PyQt5.QtWidgets import QApplication, QWidget, QLineEdit, QFileDialog, QPushButton, QLabel, QMessageBox, \
    QErrorMessage, QMainWindow
import PDFC
from PyQt5.QtCore import Qt, pyqtSlot
import PyQt5.QtCore as core
from PyQt5.QtWidgets import QVBoxLayout,QDesktopWidget
from gui_elements import stylesheets
from file_manager import *

class SpecialBG(QLabel):
    def __init__(self, parent):
        QLabel.__init__(self, parent)
        label = QLabel(self)
        pixmap = QPixmap('img/bg_img_rounded_s.png') # if 'libpng warning: iCCP: CRC error' remove png color profile
        label.setPixmap(pixmap)
        label.move(268, 150)
        label.setStyleSheet("""
            border-radius: 10;
            """)
        self.setStyleSheet(stylesheets.main_window_style)



class PDFCApp(QWidget):

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.title = 'PDFC'
        self.left = 300
        self.top = 300
        self.width = 800
        self.height = 450
        self.file_excel = ''
        self.file_list = []
        self.initUI()

##### Metodi che rendono possibile il drag&drop dei file ################

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        self.file_list = [u.toLocalFile() for u in event.mimeData().urls()]
        approved_list, rejected_list = check_if_excel_file(self.file_list)
        print(f'Approvati: {approved_list}')
        print(f'Rifiutati: {rejected_list}')

##### Metodi che rendono possibile spostare la finestra senza bordi #####

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = core.QPoint(event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()

##########################################################################


    def initUI(self):
        self.center()
        self.oldPos = self.pos()
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setAcceptDrops(True)
        VBox = QVBoxLayout()
        roundyround = SpecialBG(self)
        VBox.addWidget(roundyround)
        self.setLayout(VBox)
        # transparency cannot be set for window BG in style sheets, so...
        # self.setWindowOpacity(0.5)
        self.setWindowFlags(
            Qt.FramelessWindowHint  # hides the window controls
            # | Qt.WindowStaysOnTopHint  # forces window to top... maybe
            # | Qt.SplashScreen  # this one hides it from the task bar!
        )
        # alternative way of making base window transparent
        self.setAttribute(Qt.WA_TranslucentBackground, True)  # 100% transparent
        self.setAcceptDrops(True)
        button_get_file = QPushButton('CARICA I FILE', self)
        button_get_file.move(50, 290)
        button_get_file.setStyleSheet(stylesheets.btn_style)
        button_get_file.resize(170,50)
        button_get_file.clicked.connect(self.on_click_1)

        button_ok = QPushButton('E CONVERTILI', self)
        button_ok.move(50, 350)
        button_ok.setStyleSheet(stylesheets.btn_style)
        button_ok.resize(170,50)
        button_ok.clicked.connect(self.on_click_ok)

        button_exit = QPushButton('X', self)
        button_exit.move(720, 10)
        button_exit.setStyleSheet(stylesheets.exit_btn_style)
        button_exit.clicked.connect(QApplication.instance().quit)

        banner = QLabel(self)
        banner.setText("Tuci.it")
        banner.setAlignment(Qt.AlignCenter)
        banner.setStyleSheet(stylesheets.banner)

        drop_label = QLabel(self)
        drop_label.setGeometry(50,50,710,100)
        drop_label.setText("Trascina qui i tuoi file o premi il pulsante qui sotto per cercarli")
        drop_label.setAlignment(Qt.AlignCenter)
        drop_label.setStyleSheet(stylesheets.drop_label)

        self.show()


    @pyqtSlot()
    def on_click_1(self):
        self.open_file_name_dialog()

    def on_click_ok(self):
        # message = PDFC.handle_csv(self.file1, self.textbox.text(), self.file2)
        self.resultLabel.setText('message')
        self.resultLabel.adjustSize()

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        self.file_list, _ = QFileDialog.getOpenFileNames(self, "Scegli un file", "",
                                                   "All Files (*)", options=options)
        approved_list, rejected_list = check_if_excel_file(self.file_list)
        print(f'Approvati: {approved_list}')
        print(f'Rifiutati: {rejected_list}')

