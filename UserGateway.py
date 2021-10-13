import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials
# from openpyxl import load_workbook
# import http.client as httplib
import socket


def is_legal():
    if have_internet():
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        jason_credentials = f"C:\\Users\\{os.getlogin()}\\OneDrive - Hexagon\\Project - Koba Alignment\\api " \
                            f"tutorial\\creds.json "
        credentials = ServiceAccountCredentials.from_json_keyfile_name(jason_credentials, scope)
        client = gspread.authorize(credentials)

        sheet = client.open("Times_test").sheet1

        # data = sheet.get_all_records()
        row1 = sheet.row_values(1)
        # row2 = sheet.row_values(2)
        print(f'Answer from sheet: {row1[11]}')

        # wb = load_workbook(filename="Deps/Details/DsName.xlsx")
        # ws = wb.active
        # ws['a100'] = str("row1[11]")
        # wb.save("Deps/Details/DsName.xlsx")
        # wb.close()
        p = True
        if row1[11].lower() == 'false':
            p = False
        return p
    else:
        return True


def have_internet():
    try:
        socket.create_connection(('Google.com', 80))
        print('Socket Connection True')
        return True
    except OSError as err:
        print(f'Error from socket connection is: {err}')
        return False
