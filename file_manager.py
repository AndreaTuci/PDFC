import os
import shutil

from win32com import client

from locale import choose_locale

lang = choose_locale()


def check_if_excel_file(self):
    for file in self.file_list:
        if file.endswith(('.xlsx', '.xls', '.ods')):
            self.approved_files.append(file)
        else:
            self.rejected_files.append(file)

    update_lbl_messages(self)


def update_lbl_messages(self):
    rejected = ''
    for file in self.rejected_files:
        (_, tail) = os.path.split(file)
        rejected += f'\n{tail}'

    message = f'{lang.files_approved}: {len(self.approved_files)}\n{lang.files_rejected}: {len(self.rejected_files)}\n\n' \
              f'{f"{lang.rejected}:{rejected}" if rejected else ""}'
    self.lbl_messages.setText(message)


def convert_files(self):
    total_sheets = 0
    main_dir = os.path.dirname(os.path.realpath('__file__'))
    sheets_done = 0
    links = ['<h4 style=\"text-transform: uppercase;\">Ecco i PDF:</h4>']
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
            while True:
                try:
                    os.makedirs(new_dir)
                    break
                except:
                    shutil.rmtree(new_dir)

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
                self.ext_thred.progress_changed.emit(sheets_done)

            sheets.Close(True)
            # \"http://www.google.com\"
            (_, dir_name) = os.path.split(new_dir)
            links.append(f"<a "
                         f"style=\"text-transform: uppercase; text-decoration: none\""
                         f"href={new_dir}>"
                         f"{dir_name}</a><br>")
            # self.lbl_messages.setText(urlLink)

        except Exception as e:
            print(f'Impossibile convertire il file {file}: {e}\n\n')
        finally:
            excel.quit()
    return links