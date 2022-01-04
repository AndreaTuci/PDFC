from PyQt5.QtCore import QLocale


def choose_locale():
    locale = QLocale()
    print(f'{locale.nativeCountryName()}, {locale.nativeLanguageName()}: {locale.name()}')
    if locale.name() == 'it_IT':
        language = Italiano()
    else:
        language = Language()
    return language


class Language:
    load_files = 'LOAD FILES'
    instructions = 'Click the button or drag and drop files here'
    url_link = 'No file was converted'
    choose_files = 'Choose files'
    files_approved = 'Approved files'
    files_rejected = 'Rejected files'
    rejected = 'Rejected'
    conversion_error = 'Error in sheet n.'
    conversion_aborted = 'Unable to convert the file'
    view_result = 'View PDFs'
    deleting_old = 'Deleting old files...'

    def show_conversion_error(self, file, i, e):
        return f'{file}\n{self.conversion_error}{i}: {e}\n\n'


class Italiano(Language):
    def __init__(self):
        Language.__init__(self)
        self.load_files = 'CARICA FILE'
        self.instructions = 'Clicca il pulsante o trascina i file qui'
        self.url_link = 'Nessun file convertito'
        self.choose_files = "Scegli i file da convertire"
        self.files_approved = 'File approvati'
        self.files_rejected = 'File rifiutati'
        self.rejected = 'Rifiutati'
        self.conversion_error = 'Errore per il foglio n.'
        self.conversion_aborted = 'Impossibile convertire il file'
        self.view_result = 'Apri i PDF'
        self.deleting_old = 'Cancello i vecchi file...'

