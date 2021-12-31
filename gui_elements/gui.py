from PyQt5.QtGui import QPixmap, QCursor, QMovie
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton, QLabel, QProgressBar, QScrollArea
from PyQt5.QtCore import Qt
import PyQt5.QtCore as Core
from PyQt5.QtWidgets import QVBoxLayout, QDesktopWidget, QGraphicsDropShadowEffect
from gui_elements import stylesheets
from file_manager import *
import pythoncom
from locale import *


lang = choose_locale()


''' Classe per definire il formato della finestra '''


class ScrollLabel(QScrollArea):

    # constructor
    def __init__(self, *args, **kwargs):
        QScrollArea.__init__(self, *args, **kwargs)

        # making widget resizable
        self.setWidgetResizable(True)

        # making qwidget object
        content = QWidget(self)
        self.setWidget(content)

        # vertical box layout
        lay = QVBoxLayout(content)

        # creating label
        self.label = QLabel(content)

        # setting alignment to the text
        self.label.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        # making label multi-line
        self.label.setWordWrap(True)

        # adding label to the layout
        lay.addWidget(self.label)
        self.label.setOpenExternalLinks(True)

    # the setText method
    def setText(self, text):
        # setting text to the label
        self.label.setText(text)

class ProcessFilesThread(Core.QThread):

    def __init__(self, app):
        super().__init__()
        self.app = app

    progress_changed = Core.pyqtSignal(int)
    update_parent_msg = Core.pyqtSignal(str)

    def run(self):
        check_if_excel_file(self.app)
        url_link = lang.url_link
        if self.app.approved_files:
            try:
                pythoncom.CoInitialize()
                links = convert_files(self.app)
                url_link = ''
                for link in links:
                    url_link += link
            except Exception as e:
                print(e)
        try:
            self.update_parent_msg.emit(url_link)
        except Exception as e:
            print(e)


class RoundedWindowWithImg(QLabel):
    def __init__(self, parent):
        QLabel.__init__(self, parent)
        self.setStyleSheet(stylesheets.main_window_style)


''' Classe principale '''


class PDFCApp(QWidget):

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.title = 'PDFC'
        self.left = 200
        self.top = 200
        self.width = 300
        self.height = 648
        self.file_excel = ''
        self.file_list = []
        self.approved_files = []
        self.rejected_files = []
        self.oldPos = ''
        self.init_ui()

    def update_results_lbl(self, links):
        self.lbl_results.setText(links)
        # self.lbl_results.adjustSize()

    def on_progress_changed(self, value):
        self.progress.setValue(value)

    def reset_process_values(self):
        self.lbl_messages.setText("")
        self.lbl_results.setText("")
        self.approved_files = []
        self.rejected_files = []

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

        shadow = QGraphicsDropShadowEffect()
        shadow.setOffset(4, 4)
        self.setGraphicsEffect(shadow)
        ##### Accetta i drop ###########
        self.setAcceptDrops(True)

        ##### Per il drag della finestra
        self.center()
        self.oldPos = self.pos()

        ##### Elementi funzionali ######

        lbl_header = QLabel(self)
        header_img = QPixmap('img/header_s.png')
        # if 'libpng warning: iCCP: CRC error' remove png color profile
        lbl_header.setPixmap(header_img)
        lbl_header.move(40, 50)
        lbl_header.setStyleSheet("""
                    border-radius: 10px;
                    """)
        lbl_header.adjustSize()

        self.progress = QProgressBar(self)
        self.progress.setGeometry(40, 170, 253, 25)
        self.progress.setMaximum(100)
        # self.progress.setAlignment(Qt.AlignCenter)
        self.progress.setFormat("")
        self.progress.hide()

        self.btn_get_files = QPushButton(lang.load_files, self)
        self.btn_get_files.move((self.width/2)-70, 75)
        self.btn_get_files.setStyleSheet(stylesheets.btn_style)
        self.btn_get_files.adjustSize()
        self.btn_get_files.clicked.connect(self.get_files)
        self.btn_get_files.installEventFilter(self)

        self.btn_exit = QPushButton('X', self)
        self.btn_exit.move(self.width-40, 20)
        self.btn_exit.setStyleSheet(stylesheets.exit_btn_style)
        self.btn_exit.clicked.connect(QApplication.instance().quit)
        self.btn_exit.installEventFilter(self)
        self.btn_exit.adjustSize()

        lbl_banner = QLabel(self)
        lbl_banner.setOpenExternalLinks(True)
        banner_link = "<a style=\"color:white; text-decoration: none\" " \
                      "href=\"https://github.com/AndreaTuci/PDFC\">" \
                      "code on GitHub</a>"
        lbl_banner.setText(banner_link)
        lbl_banner.setAlignment(Qt.AlignCenter)
        lbl_banner.setStyleSheet(stylesheets.banner)
        lbl_banner.adjustSize()

        self.lbl_messages = ScrollLabel(self)
        self.lbl_messages.setGeometry(43, 200, 214, 150)
        self.lbl_messages.setText("")
        # self.lbl_messages.setAlignment(Qt.AlignLeft)
        self.lbl_messages.setStyleSheet(stylesheets.message_label)

        self.lbl_results = ScrollLabel(self)
        self.lbl_results.setGeometry(43, 400, 214, 200)
        self.lbl_results.setText("")
        # self.lbl_results.setAlignment(Qt.AlignLeft)
        self.lbl_results.setStyleSheet(stylesheets.message_label)
        # self.lbl_results.setOpenExternalLinks(True)
        self.show()


    ##### Metodi che rendono possibile il drag&drop dei file ################

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        self.file_list = [u.toLocalFile() for u in event.mimeData().urls()]
        self.start_converting_process()



    def start_converting_process(self):
        self.progress.show()
        self.reset_process_values()
        self.ext_thred = ProcessFilesThread(app=self)
        self.ext_thred.progress_changed.connect(self.on_progress_changed)
        self.ext_thred.update_parent_msg.connect(self.update_results_lbl)
        self.ext_thred.start()


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

    def eventFilter(self, obj, event):
        if obj == self.btn_get_files and event.type() == Core.QEvent.HoverEnter:
            self.on_hovered(obj)
        elif obj == self.btn_get_files and event.type() == Core.QEvent.Leave:
            self.leaveEvent(event)
        elif obj == self.btn_exit and event.type() == Core.QEvent.HoverEnter:
            self.on_hovered(obj)
        return super(PDFCApp, self).eventFilter(obj, event)

    def on_hovered(self, obj):
        obj.setCursor(QCursor(Qt.PointingHandCursor))
        if obj == self.btn_get_files:
            self.btn_get_files.setStyleSheet(stylesheets.btn_style_hover)

    def leaveEvent(self, e):
        self.btn_get_files.setStyleSheet(stylesheets.btn_style)

    def get_files(self):
        self.file_list, _ = QFileDialog.getOpenFileNames(self, lang.choose_files, "", "All Files (*)")
        self.start_converting_process()
