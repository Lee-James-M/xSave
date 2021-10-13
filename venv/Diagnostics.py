import os
import datetime
import tkinter.messagebox

# map config of 104 check again atreport/serv/
# 2009 against uncertainty
interrogation_data = []

def main_diagnostics():
    interrogation_data.clear()
    interrogation_data.append(f'Hi {os.getlogin().split(".")[0]},')
    interrogation_data.append(f'''Your Diagnostic System Assessment on {datetime.datetime.now().strftime('%d/%m/%Y')} at {datetime.datetime.now().strftime("%H:%M:%S")}.''')
    interrogation_data.append("---------------------------\nCurrent calibration configuration machine data:")
    # get machine data
    if os.path.exists("C:/PCDMISW/PCDDATA/IsoMac.dat"):
        print("FILE EXISTS")
        with open("C:/PCDMISW/PCDDATA/IsoMac.dat") as f:
            _data = f.readlines()
            print("isomac" + str(_data))
            interrogation_data.append("Machine:    " + _data[1].strip('\n'))
            interrogation_data.append("size:       " + _data[3].strip('\n'))
            interrogation_data.append("Controller: " + _data[7].strip('\n'))
    else:
        interrogation_data.append("Machine data not available")
    interrogation_data.append("---------------------------")

def setup_diagnostics():
    find_list_autotune = ["Autotune", "AutotuneDC", "autotunedc"]
    for item in os.listdir("C:/"):
        for _file in find_list_autotune:
            if _file == item:
                item = item.strip("\n")
                interrogation_data.append(f"Detected: {item}")
    find_list_pcdmis = ["2018", "2019", "2020", "2021"]
    for item in os.listdir("C:/Program Files/Hexagon/"):
        for _file in find_list_pcdmis:
            if _file in item:
                if "Help" not in item:
                    interrogation_data.append(f"Detected: {item}")
    interrogation_data.append("---------------------------")
    for item in os.listdir("C:\\Program Files\\Hexagon\\"):
        _dir = os.listdir("C:\\Program Files\\Hexagon\\" + item)
        if "interfac.dll" in _dir:
            if os.path.getsize("C:\\Program Files\\Hexagon\\" + item + "/interfac.dll") == os.path.getsize(
                    "C:\\Program Files\\Hexagon\\" + item + "/leitz.dll"):
                interrogation_data.append(f"{item} version of PC-DMIS is configured with an leitz.dll interfac")
            if os.path.getsize("C:\\Program Files\\Hexagon\\" + item + "/interfac.dll") == os.path.getsize(
                    "C:\\Program Files\\Hexagon\\" + item + "/fdc.dll"):
                interrogation_data.append(f"{item} version of PC-DMIS is configured with an fdc.dll interfac")
    interrogation_data.append("---------------------------")

def uncertainty_diagnostics():
    if os.path.exists("C:/Program Files (x86)/DeaReport/ATReport.ini"):
        with open("C:/Program Files (x86)/DeaReport/ATReport.ini") as f:
            _data = f.readlines()
            print(_data)
            _data = _data[2].strip("\n")
            print(f"data: {_data}")
            if _data == "DisabilitaIncertezza=1":
                interrogation_data.append("Uncertainty: Disabled")
            else:
                interrogation_data.append("Uncertainty: Enabled")
    else:
        interrogation_data.append("Uncertainty data not available")
    interrogation_data.append("---------------------------")

def serv_diagnostics(cal_config):
    atdir = os.listdir(cal_config.get_autotune_dir())
    for item in atdir:
        if item.lower() == "serv.stp":
            with open(os.path.join(cal_config.get_autotune_dir(), item)) as f:
                for i in range(1, 8):
                    sizes = f.readline()
                for i in range(1, 9):
                    f.readline()
                line_sixteen = f.readline()
                line_sixteen = line_sixteen.strip('\n')
                line_seventeen = f.readline()
                line_seventeen = line_seventeen.strip('\n')
                sizes = str(sizes).strip('\n')
                sizes = sizes.replace('.0', '')
                sizes = sizes.split(' ')
                interrogation_data.append(f'Machine data from uploaded serv file from: '
                                          f'{cal_config.get_autotune_dir()}\nLyy: {sizes[2]}\nLxx: {sizes[1]}'
                                          f'\nLzz: {sizes[3]}\nGranite: {line_sixteen}\nBridge:  {line_seventeen}')

def build_diagnsotic_report():
    interrogation_data.append("---------------------------\n---------------------------\nRecommendations:")
    saved_diagnostic_report = f'''C:/Users/{os.getlogin()}/dataAssessment{datetime.datetime.now().strftime('_%H.%M.%S_%d-%m-%Y')}.txt'''
    f = open(saved_diagnostic_report, "w+")
    for item in interrogation_data:
        f.write(item + "\n")
    os.startfile(saved_diagnostic_report)

def run_diagnostics(cal_config):
    main_diagnostics()
    setup_diagnostics()
    uncertainty_diagnostics()
    serv_diagnostics(cal_config)
    build_diagnsotic_report()

