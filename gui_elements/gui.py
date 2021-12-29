from PyQt5.QtGui import QPixmap, QCursor, QMovie
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton, QLabel, QProgressBar
from PyQt5.QtCore import Qt, pyqtSlot
import PyQt5.QtCore as Core
from PyQt5.QtWidgets import QVBoxLayout, QDesktopWidget, QGraphicsDropShadowEffect
from gui_elements import stylesheets
from file_manager import *
import pythoncom


''' Classe per definire il formato della finestra '''

class External(Core.QThread):

    def __init__(self, app):
        super().__init__()
        self.app = app


    countChanged = Core.pyqtSignal(int)

    def run(self):
        check_if_excel_file(self.app)
        try:
            pythoncom.CoInitialize()

            links = convert_files(self.app)
            urlLink = ''
            for link in links:
                urlLink += link
            # self.app.lbl_messages.setText(urlLink)
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
        self.left = 300
        self.top = 300
        self.width = 800
        self.height = 450
        self.file_excel = ''
        self.file_list = []
        self.approved_files = []
        self.rejected_files = []
        self.oldPos = ''
        self.movie = QMovie("img/loading.gif")
        self.init_ui()

    def onButtonClick(self):
        return

    def onCountChanged(self, value):
        self.progress.setValue(value)

    def init_ui(self):


        # self.startAnimation()

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

        self.progress = QProgressBar(self)
        self.progress.setGeometry(50, 170, 253, 25)
        self.progress.setMaximum(100)

        lbl_header = QLabel(self)
        header_img = QPixmap('img/header_s.png')
        # if 'libpng warning: iCCP: CRC error' remove png color profile
        lbl_header.setPixmap(header_img)
        lbl_header.setGeometry(50, 50, 112, 43)
        lbl_header.setStyleSheet("""
                    border-radius: 10px;
                    """)
        lbl_header.adjustSize()

        self.btn_get_files = QPushButton('+', self)
        self.btn_get_files.move(324, 68)
        self.btn_get_files.setStyleSheet(stylesheets.btn_style)
        self.btn_get_files.resize(50, 50)
        self.btn_get_files.clicked.connect(self.get_files)
        self.btn_get_files.installEventFilter(self)

        self.btn_exit = QPushButton('X', self)
        self.btn_exit.move(760, 20)
        self.btn_exit.setStyleSheet(stylesheets.exit_btn_style)
        self.btn_exit.clicked.connect(QApplication.instance().quit)
        self.btn_exit.installEventFilter(self)
        self.btn_exit.adjustSize()

        lbl_banner = QLabel(self)
        lbl_banner.setOpenExternalLinks(True)
        banner_link = "<a style=\"color:white; text-decoration: none\" href=\"https://github.com/AndreaTuci/PDFC\">code on GitHub</a>"
        lbl_banner.setText(banner_link)
        lbl_banner.setAlignment(Qt.AlignCenter)
        lbl_banner.setStyleSheet(stylesheets.banner)
        lbl_banner.adjustSize()

        self.lbl_loading_icon = QLabel(self)
        self.lbl_loading_icon.move(200, 200)
        self.lbl_loading_icon.setMovie(self.movie)
        self.lbl_loading_icon.adjustSize()

        self.lbl_messages = QLabel(self)
        self.lbl_messages.setGeometry(53, 200, 100, 130)
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
        self.calc = External(app=self)
        self.calc.countChanged.connect(self.onCountChanged)
        self.calc.start()



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
            self.onHovered(obj)
        elif obj == self.btn_get_files and event.type() == Core.QEvent.Leave:
            self.leaveEvent(event)
        elif obj == self.btn_exit and event.type() == Core.QEvent.HoverEnter:
            self.onHovered(obj)
        return super(PDFCApp, self).eventFilter(obj, event)

    def onHovered(self, obj):
        obj.setCursor(QCursor(Qt.PointingHandCursor))
        if obj == self.btn_get_files:
            self.btn_get_files.setStyleSheet(stylesheets.btn_style_hover)

    def leaveEvent(self, e):
        self.btn_get_files.setStyleSheet(stylesheets.btn_style)

    def startAnimation(self):
        self.movie.start()

    def stopAnimation(self):
        self.movie.stop()

    @pyqtSlot()
    def get_files(self):
        options = QFileDialog.Options()
        self.file_list, _ = QFileDialog.getOpenFileNames(self, "Scegli i file da convertire", "", "All Files (*)", options=options)
        check_if_excel_file(self)
        convert_files(self)






