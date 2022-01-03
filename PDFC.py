import sys
from gui_elements import gui
from locales import choose_locale

lang = choose_locale()

if __name__ == '__main__':
    app = gui.QApplication(sys.argv)
    application = gui.PDFCApp()
    sys.exit(app.exec_())
