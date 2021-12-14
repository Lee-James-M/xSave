import xml.etree.ElementTree as eT  # for autotuneDC xml params file


params_files = [('C:\\AutotuneDC\\paramX.xml', 'X'), ('C:\\AutotuneDC\\paramY.xml', 'Y'), ('C:\\AutotuneDC\\paramZ.xml',
                                                                                           'Z')]

adj_files = []

for file in params_files:
    tree = eT.parse(file[0])
    adj_number = tree.findall('adjFile/Value')
    adj_files.append(f'{file[1]}{adj_number[0].text}.adj')

for file in adj_files:
    print(file)
