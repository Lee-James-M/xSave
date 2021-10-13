import os

dependencies_location = f'C:/Users/{os.getlogin()}/OneDrive - Hexagon/Deps_xSave/'

# LISTS
pcdmis_versions = ['PC-DMIS 2021.2 64-bit', 'PC-DMIS 2021.1 64-bit', 'PC-DMIS 2020 R2 64-bit', 'PC-DMIS 2020 R1 64-bit',
                   'PC-DMIS 2019 R2 64-bit']

autotune_ext_delete_list = ["adjtest1.txt", "adjtest2.txt", "adjtest3.txt", "autadj.log", "constant.asc",
                            "creacos.asc", "deacfg.asc", "e2pot.dat", "e2potscl.dat", "espot.dat", "espotscl.dat",
                            "pcdparams.txt", "sensor.asc", "siqfiles.tab", "constant.edt,", ".adj", ".stp"]

autotunedc_ext_delete_list = ["adjtest1.txt", "adjtest2.txt", "adjtest3.txt", "autadj.log", "deacfg.asc",
                              "paramcommon.xml", "paramhw.xml", "paramsys.xml", "paramx.xml", "paramy.xml",
                              "paramz.xml", "pcdparams.txt", "prbparams.txt"]

machine_adj_list = (
    ('global a', 'GLOA'),
    ('global b', 'GLOB'),
    ('global c', 'GLOC'),
    ('global d', 'GLOD'),
    ('global e', 'GLOE'),
    ('global f', 'GLOF'),
    ('tigo', 'TIGO'),
    ('pioneer', 'PIO'),
    ('micra', 'MIC'),
    ('alpha', 'ALPHA'),
    ('delta', 'DELTA'),
    ('xcel', 'XCEL'),
    ('chameleon', 'CHAM'),
    ('vento', 'VENTO'),
    ('bravo', 'BRAVO'),
    ('oxalis', 'OXALIS'),
    ('lk', 'LK'),
    ('gamma', 'GAMMA'),
    ('sirrocco', 'SIRROCCO'),
    ('mistral', 'MIST')
)
machine_list = [
    "Global AA", "Global A", "Global B", "Global C", "Global D", "Global E", "Global F", "Micra", "Tigo", "Pioneer",
    "Alpha", "Delta", "Xcel", "Chameleon", "Vento", "Oxalis", "Mistral", "Bravo", "LK", "Gamma", "Sirrocco",
]
# "kobareport.dat" - removed
pcddata_delete_list = ["isoblo.dat", "isocmm.dat", "isocust.dat", "isokoba.dat", "isomac.dat", "isoscan.dat",
                       "isosfe.dat", "isotol.dat", "isout.dat"]
