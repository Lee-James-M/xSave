import os.path
import shutil
import xml.etree.ElementTree as eT  # for autotuneDC xml params file
import tkinter.messagebox


# only for DC so far
def get_dc_adj_files():
    params_files = [('C:\\AutotuneDC\\paramX.xml', 'X'), ('C:\\AutotuneDC\\paramY.xml', 'Y'),
                    ('C:\\AutotuneDC\\paramZ.xml', 'Z')]
    adj_files = []

    for file in params_files:
        tree = eT.parse(file[0])
        adj_number = tree.findall('adjFile/Value')
        adj_files.append(f'{file[1]}{adj_number[0].text}.adj')

    print(adj_files)

    tree = eT.parse('C:\\AutotuneDC\\paramCommon.xml')
    serial_number = tree.findall('serialNumber/Value')
    print(serial_number[0].text)

    return adj_files


get_dc_adj_files()

try:
    for adj in get_dc_adj_files():
        shutil.copyfile(os.path.join('C:/AutotuneDC/XSMCS 1815 AUTOTUNE DATA DC V 05_7', adj),
                        os.path.join('c:/adj', adj))

except FileNotFoundError as e:
    print(f'{e}\n\nUnable to complete transfer of ADJ files as one or more uploaded machine files was not found.')
    tkinter.messagebox.showerror('File not found', f'Unable to complete transfer of ADJ files as one or more '
                                                   f'uploaded machine files was not found.\n\n{e}')
