from PyQt5.QtGui import QPixmap, QCursor
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton, QLabel, QProgressBar, QScrollArea
from PyQt5.QtCore import Qt
import PyQt5.QtCore as Core
from PyQt5.QtWidgets import QVBoxLayout, QDesktopWidget, QGraphicsDropShadowEffect
from gui_elements import stylesheets
from file_manager import *
import pythoncom
from PDFC import lang


''' Secondary classes '''

# Makes a label scrollable


class ScrollLabel(QScrollArea):

    def __init__(self, *args, **kwargs):
        QScrollArea.__init__(self, *args, **kwargs)
        self.setWidgetResizable(True)
        content = QWidget(self)
        self.setWidget(content)
        layout = QVBoxLayout(content)
        # Declaring Label into QScrollArea does the trick
        self.label = QLabel(content)
        self.label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.label.setWordWrap(True)
        layout.addWidget(self.label)
        self.label.setOpenExternalLinks(True)

    def setText(self, text):
        # Need to override setText method
        self.label.setText(text)


# A secondary thread allows to process files while updating the process bar

class ProcessFilesThread(Core.QThread):

    def __init__(self, app):
        super().__init__()
        self.app = app

    # These are the signals we use to communicate between threads
    progress_changed = Core.pyqtSignal(int)
    update_parent_msg = Core.pyqtSignal(str)

    def run(self):
        check_if_excel_file(self.app)
        url_link = lang.url_link
        if self.app.approved_files:
            try:
                # CoInitialize needed to make the Excel thread in convert_files work
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


# Class used to create rounded corners for the main window

class RoundedWindow(QLabel):
    def __init__(self, parent):
        QLabel.__init__(self, parent)
        self.setStyleSheet(stylesheets.main_window_style)


''' Main class '''


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

    def init_ui(self):

        # Basic GUI settings ###########

        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setAcceptDrops(True)
        v_box = QVBoxLayout()
        window_frame = RoundedWindow(self)
        v_box.addWidget(window_frame)
        self.setLayout(v_box)
        self.setWindowFlags(
            Qt.FramelessWindowHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        shadow = QGraphicsDropShadowEffect()
        shadow.setOffset(4, 4)
        self.setGraphicsEffect(shadow)

        lbl_header = QLabel(self)
        application_path = os.path.dirname(os.path.realpath('__file__'))
        header_path = os.path.join(application_path, "img", "header_s.png")
        header_img = QPixmap(header_path)
        # if 'libpng warning: iCCP: CRC error' remove png color profile
        lbl_header.setPixmap(header_img)
        lbl_header.move(40, 50)
        lbl_header.setStyleSheet("""
                            border-radius: 10px;
                            """)
        lbl_header.adjustSize()

        lbl_banner = QLabel(self)
        lbl_banner.setOpenExternalLinks(True)
        banner_link = "<a style=\"color:white; text-decoration: none\" " \
                      "href=\"https://github.com/AndreaTuci/PDFC\">" \
                      "code on GitHub</a>"
        lbl_banner.setText(banner_link)
        lbl_banner.setAlignment(Qt.AlignCenter)
        lbl_banner.setStyleSheet(stylesheets.banner)
        lbl_banner.adjustSize()

        # Enables drop files ###########

        self.setAcceptDrops(True)

        # Enables window dragging ###########

        self.center()
        self.oldPos = self.pos()

        # Functional elements ###########

        self.progress = QProgressBar(self)
        self.progress.setGeometry(40, 170, 253, 25)
        self.progress.setMaximum(100)
        self.progress.setFormat("")
        self.progress.hide()

        self.btn_get_files = QPushButton(lang.load_files, self)
        self.btn_get_files.move((self.width/2)-70, 75)
        self.btn_get_files.setStyleSheet(stylesheets.btn_style)
        self.btn_get_files.adjustSize()
        self.btn_get_files.clicked.connect(self.get_files)
        self.btn_get_files.installEventFilter(self)

        self.lbl_messages = ScrollLabel(self)
        self.lbl_messages.setGeometry(43, 200, 214, 150)
        self.lbl_messages.setText("")
        self.lbl_messages.setStyleSheet(stylesheets.message_label)

        self.lbl_results = ScrollLabel(self)
        self.lbl_results.setGeometry(43, 400, 214, 200)
        self.lbl_results.setText("")
        self.lbl_results.setStyleSheet(stylesheets.message_label)

        self.btn_exit = QPushButton('X', self)
        self.btn_exit.move(self.width - 40, 20)
        self.btn_exit.setStyleSheet(stylesheets.exit_btn_style)
        self.btn_exit.clicked.connect(QApplication.instance().quit)
        self.btn_exit.installEventFilter(self)
        self.btn_exit.adjustSize()

        self.show()

    # Methods for files drag & drop ###########

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        self.file_list = [u.toLocalFile() for u in event.mimeData().urls()]
        self.start_converting_process()

    # Methods used for window drag ###########

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

    # This handle hovering over buttons, changing their style ###########

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

    # Functional methods ###########

    def get_files(self):
        self.file_list, _ = QFileDialog.getOpenFileNames(self, lang.choose_files, "", "All Files (*)")
        self.start_converting_process()

    def start_converting_process(self):
        self.progress.show()
        self.reset_process_values()

        # Starting the second thread

        self.ext_thred = ProcessFilesThread(app=self)
        self.ext_thred.progress_changed.connect(self.on_progress_changed)
        self.ext_thred.update_parent_msg.connect(self.update_results_lbl)
        self.ext_thred.start()

    def update_results_lbl(self, links):
        self.lbl_results.setText(links)

    def on_progress_changed(self, value):
        self.progress.setValue(value)

    def reset_process_values(self):
        self.lbl_messages.setText("")
        self.lbl_results.setText("")
        self.approved_files = []
        self.rejected_files = []
