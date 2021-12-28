from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton, QLabel
from PyQt5.QtCore import Qt, pyqtSlot
import PyQt5.QtCore as Core
from PyQt5.QtWidgets import QVBoxLayout, QDesktopWidget

import PDFC
from gui_elements import stylesheets
from file_manager import *


''' Classe per definire il formato della finestra '''


class RoundedWindowWithImg(QLabel):
    def __init__(self, parent):
        QLabel.__init__(self, parent)
        label = QLabel(self)
        pixmap = QPixmap('img/bg_img_rounded_s.png')
        # if 'libpng warning: iCCP: CRC error' remove png color profile
        label.setPixmap(pixmap)
        label.move(268, 150)
        label.setStyleSheet("""
            border-radius: 10;
            """)
        self.setStyleSheet(stylesheets.main_window_style)


''' Classe principale '''


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
        self.approved_files = []
        self.rejected_files = []
        self.oldPos = ''
        self.init_ui()

    def init_ui(self):

        ##### Impostazioni base

        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setAcceptDrops(True)
        v_box = QVBoxLayout()
        window_frame = RoundedWindowWithImg(self)
        v_box.addWidget(window_frame)
        self.setLayout(v_box)
        self.setWindowFlags(
            Qt.FramelessWindowHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        ##### Accetta i drop ###########
        self.setAcceptDrops(True)

        ##### Per il drag della finestra
        self.center()
        self.oldPos = self.pos()

        ##### Elementi funzionali ######

        btn_get_files = QPushButton('CARICA I FILE', self)
        btn_get_files.move(50, 290)
        btn_get_files.setStyleSheet(stylesheets.btn_style)
        btn_get_files.resize(170, 50)
        btn_get_files.clicked.connect(self.get_files)

        btn_convert = QPushButton('E CONVERTILI', self)
        btn_convert.move(50, 350)
        btn_convert.setStyleSheet(stylesheets.btn_style)
        btn_convert.resize(170, 50)
        btn_convert.clicked.connect(self.convert_files)

        btn_exit = QPushButton('X', self)
        btn_exit.move(720, 10)
        btn_exit.setStyleSheet(stylesheets.exit_btn_style)
        btn_exit.clicked.connect(QApplication.instance().quit)

        lbl_banner = QLabel(self)
        lbl_banner.setText("Tuci.it")
        lbl_banner.setAlignment(Qt.AlignCenter)
        lbl_banner.setStyleSheet(stylesheets.banner)

        lbl_drop_files = QLabel(self)
        lbl_drop_files.setGeometry(50, 50, 710, 100)
        lbl_drop_files.setText("Trascina qui i tuoi file o premi il pulsante qui sotto per caricarli")
        lbl_drop_files.setAlignment(Qt.AlignCenter)
        lbl_drop_files.setStyleSheet(stylesheets.drop_label)

        self.lbl_messages = QLabel(self)
        self.lbl_messages.setGeometry(53, 170, 100, 130)
        self.lbl_messages.setText("")
        self.lbl_messages.setAlignment(Qt.AlignLeft)
        self.lbl_messages.setStyleSheet(stylesheets.message_label)
        self.lbl_messages.setOpenExternalLinks(True)


        self.show()

    ##### Metodi che rendono possibile il drag&drop dei file ################

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        self.file_list = [u.toLocalFile() for u in event.mimeData().urls()]
        check_if_excel_file(self)


    ##### Metodi che rendono possibile spostare la finestra senza bordi #####

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = Core.QPoint(event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()

##########################################################################

    @pyqtSlot()
    def get_files(self):
        self.open_file_name_dialog()

    def convert_files(self):
        convert_files(self)

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        self.file_list, _ = QFileDialog.getOpenFileNames(self, "Scegli un file", "", "All Files (*)", options=options)
        check_if_excel_file(self)


