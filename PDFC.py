import sys
from gui_elements import gui



if __name__ == '__main__':
    app = gui.QApplication(sys.argv)
    application = gui.PDFCApp()
    sys.exit(app.exec_())

