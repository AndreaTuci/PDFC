import os
import shutil
from win32com import client

from PDFC import lang


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

    message = f'{lang.files_approved}: {len(self.approved_files)}\n{lang.files_rejected}: ' \
              f'{len(self.rejected_files)}\n\n' \
              f'{f"{lang.rejected}:{rejected}" if rejected else ""}'
    self.lbl_messages.setText(message)


def convert_files(self):
    total_sheets = 0
    main_dir = os.path.dirname(os.path.realpath('__file__'))
    sheets_done = 0
    links = [f'<p style=\"text-transform: uppercase;\">{lang.view_result}:</p>']
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
                except Exception as e:
                    print(f'{e}\n{lang.deleting_old}')
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
                    error = str(e.excepinfo[2])
                    print(f'{lang.show_conversion_error(file, i, error)}')
                i += 1
                sheets_done += 1
                self.ext_thred.progress_changed.emit(sheets_done)

            sheets.Close(True)
            (_, dir_name) = os.path.split(new_dir)
            links.append(f"<a "
                         f"style=\"text-transform: uppercase; text-decoration: none\""
                         f"href={new_dir}>"
                         f"{dir_name}</a><br>")
            # self.lbl_messages.setText(urlLink)

        except Exception as e:
            print(f'{lang.conversion_aborted} {file}: {e}\n\n')
        finally:
            excel.quit()
    return links
