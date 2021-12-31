from PyQt5.QtCore import QLocale

def choose_locale():
    locale = QLocale()
    print(locale.nativeCountryName(), locale.name(), locale.nativeLanguageName())
    if locale.name() == 'it_IT':
        language = Italiano
    else:
        language = English
    return language


class Italiano():
    load_files = 'CARICA FILE'
    url_link = 'Nessun file convertito'
    choose_files = "Scegli i file da convertire"
    files_approved = 'File approvati'
    files_rejected = 'File rifiutati'
    rejected = 'Rifiutati'

class English():
    load_files = 'LOAD FILES'