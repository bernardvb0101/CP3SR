from openpyxl import Workbook

def write_to_xls(content, type):
    """
    This function writes a list of dictionaries that gets passed to it, to an Excel file
    """
    if type == 'url':
        new_xlxs_file_name = "./Variables/CP3_sites.xlsx"
    else:
        new_xlxs_file_name = "./Variables/CP3_API_creds.xlsx"
    wb = Workbook(write_only=True)
    sheet = wb.create_sheet()
    headers = list(content[0].keys())
    sheet.append(headers)

    for x in content:
        sheet.append(list(x.values()))

    wb.save(new_xlxs_file_name)

    return new_xlxs_file_name