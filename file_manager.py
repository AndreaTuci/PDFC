import os
from win32com import client

def check_if_excel_file(self):
    for file in self.file_list:
        if file.endswith(('.xlsx', '.xls', '.ods')):
            self.approved_files.append(file)
        else:
            self.rejected_files.append(file)

    update_lbl_messages(self)


def update_lbl_messages(self):
    print(f'Prima {self.rejected_files}')
    rejected = ''
    for file in self.rejected_files[0:3]:
        (_, tail) = os.path.split(file)
        rejected += f'\n{tail}'
    if len(self.rejected_files) > 3:
        rejected += '\n[...]'
    message = f'File approvati: {len(self.approved_files)}\nFile rifiutati: {len(self.rejected_files)}\n\n' \
              f'{f"Rifiutati:{rejected}" if rejected else ""}'
    self.lbl_messages.setText(message)


def convert_files(self):
    total_sheets = 0
    main_dir = os.path.dirname(os.path.realpath('__file__'))
    sheets_done = 0
    links = []
    for file in self.approved_files:
        try:
            excel = client.Dispatch("Excel.Application")
            excel_file_path = os.path.join(main_dir, file)
            sheets = excel.Workbooks.Open(excel_file_path)
            total_sheets += len(sheets.Worksheets)
            excel.quit()
        except Exception as e:
            print(e)

    self.progress.setMaximum(total_sheets)
    for file in self.approved_files:
        excel = client.Dispatch("Excel.Application")
        try:
            excel_file_path = os.path.join(main_dir, file)

            if file.endswith('.xlsx'):
                new_dir = file[:-5]
            else:
                new_dir = file[:-4]

            new_dir = os.path.join(main_dir, new_dir)
            os.makedirs(new_dir)

            sheets = excel.Workbooks.Open(excel_file_path)

            i = 0

            while i < len(sheets.Worksheets):
                try:
                    work_sheet = sheets.Worksheets[i]
                    title = work_sheet.Name.replace(" ", "_")
                    pdf_file_path = os.path.join(new_dir, f'{title}.pdf')
                    work_sheet.ExportAsFixedFormat(0, pdf_file_path)
                except Exception as e:
                    e = e.excepinfo[2]
                    print(f'File {file}\nErrore per il foglio n.{i}: {e}\n\n')
                i += 1
                sheets_done += 1
                self.calc.countChanged.emit(sheets_done)

            sheets.Close(True)

            links.append("<a href=\"http://www.google.com\">'Click this link to go to Google'</a>")
            # self.lbl_messages.setText(urlLink)
        except:
            print(f'Impossibile convertire il file {file}\n\n')
        finally:
            excel.quit()

    return links