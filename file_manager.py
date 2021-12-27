

def check_if_excel_file(file_list):
    print(file_list)
    approved_list = []
    rejected_list = []
    for file in file_list:
        if file.endswith(('.xlsx', '.xls', '.ods')):
            approved_list.append(file)
        else:
            rejected_list.append(file)
    return approved_list, rejected_list