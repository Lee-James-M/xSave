import os
import xml.etree.ElementTree as eT
from datetime import datetime as dt


class Diagnostics:
    def __init__(self, cal_config, model, size, serial_no, controller, temp_comp_a, temp_comp_b, atreport_temp_comp,
                 serv_lxx, serv_lyy, serv_lzz, serv_line_sixteen, serv_line_seventeen):
        self.cal_config = cal_config
        self.model = model
        self.size = size
        self.serial_no = serial_no
        self.controller = controller
        self.temp_comp_a = temp_comp_a
        self.temp_comp_b = temp_comp_b
        self.temp_comp = ""
        self.atreport_temp_comp = atreport_temp_comp
        self.serv_lxx_stroke = serv_lxx
        self.serv_lyy_stroke = serv_lyy
        self.serv_lzz_stroke = serv_lzz
        self.serv_line_sixteen = serv_line_sixteen
        self.serv_line_sixteen_values = ''
        self.serv_line_seventeen_values = ''
        self.serv_line_seventeen = serv_line_seventeen
        self.serv_file_exists = False
        self.isomac_exists = False
        self.uncertainty_exists = False
        self.serv_stroke_equality = ""
        self.is_uncertainty = ""
        self.probe = ''
        self.gcomp_temp_comp = ''
        self.alert = []

    def show_diagnostics_class_data(self):
        # print(f"model: {self.model}\nsize: {self.size}\nserial no: {self.serial_no}\ncontroller: {self.controller}"
        #       f"\ntemp_comp_a: {self.temp_comp_a}\ntemp_comp_b: {self.temp_comp_b}\ntemp comp: {self.temp_comp}")
        self.alert.append(
            f"model: {self.model}\nsize: {self.size}\nserial no: {self.serial_no}\ncontroller: {self.controller}"
            f"\ntemp_comp_a: {self.temp_comp_a}\ntemp_comp_b: {self.temp_comp_b}\ntemp comp: {self.temp_comp}")

    # needed?
    def clear_diagnostics_class_data(self):
        # todo clear out the diag object.
        self.model = 'empty'
        self.size = ''
        self.serial_no = 0
        self.controller = ''
        self.temp_comp_a = 0
        self.temp_comp_b = 0
        self.temp_comp = ""
        self.atreport_temp_comp = ''
        self.serv_lxx_stroke = ''
        self.serv_lyy_stroke = ''
        self.serv_lzz_stroke = ''
        self.serv_line_sixteen = ''
        self.serv_line_sixteen_values = 0
        self.serv_line_seventeen_values = 0
        self.serv_line_seventeen = ''
        self.serv_file_exists = False
        self.isomac_exists = False
        self.uncertainty_exists = False
        self.serv_stroke_equality = ''
        self.is_uncertainty = ''
        self.probe = ''
        self.gcomp_temp_comp = ''
        self.alert = []

    def detect_pcdmis_versions(self):
        find_list_pcdmis = ["2017", "2018", "2019", "2020", "2021", "2022"]
        for item in os.listdir("C:/Program Files/Hexagon/"):
            for _file in find_list_pcdmis:
                if _file in item:
                    if "Help" not in item:
                        print(f"Detected: {item}")
                        self.alert.append(f"Detected: {item}")

    def detect_pcdmis_interfacs(self):
        for item in os.listdir("C:\\Program Files\\Hexagon\\"):
            _dir = os.listdir("C:\\Program Files\\Hexagon\\" + item)
            if "interfac.dll" in _dir:
                if os.path.getsize("C:\\Program Files\\Hexagon\\" + item + "/interfac.dll") == os.path.getsize(
                        "C:\\Program Files\\Hexagon\\" + item + "/leitz.dll"):
                    print(f"{item} version of PC-DMIS is configured with a leitz.dll interfac")
                    self.alert.append(f"{item} version of PC-DMIS is configured with a leitz.dll interfac")
                if os.path.getsize("C:\\Program Files\\Hexagon\\" + item + "/interfac.dll") == os.path.getsize(
                        "C:\\Program Files\\Hexagon\\" + item + "/fdc.dll"):
                    print(f"{item} version of PC-DMIS is configured with an fdc.dll interfac")
                    self.alert.append(f"{item} version of PC-DMIS is configured with an fdc.dll interfac")

    # maybe check .exe and other file are present along with the path

    def detect_autotune_version(self):
        find_list_autotune = ["Autotune", "AutotuneDC", "autotunedc"]
        for item in os.listdir("C:/"):
            for _file in find_list_autotune:
                if _file == item:
                    item = item.strip("\n")
                    print(f"Detected: {item}")
                    self.alert.append(f"Detected: {item}")

    def detect_autotune_files_exist(self):
        if os.path.exists(os.path.join(self.cal_config.get_autotune_dir(), 'paramCommon.xml')) or \
                os.path.exists(os.path.join(self.cal_config.get_autotune_dir(), 'CONSTANT.ASC')):
            return True
        else:
            return False

    # Get data from last generate isomac.dat ATReport data
    def get_isomac(self):
        if os.path.exists("C:\\pcdmisw\\pcddata\\isomac.dat"):
            with open("C:\\pcdmisw\\pcddata\\isomac.dat", "r") as f:
                isomac = f.readlines()
                # isomac.dat data in a list
                self.model = isomac[1].strip("\n")
                self.size = isomac[3].strip("\n")
                self.serial_no = isomac[4].strip("\n")
                self.controller = isomac[7].strip("\n")
                self.temp_comp_a = isomac[8].strip("\n")
                self.temp_comp_b = isomac[9].strip("\n")
                if int(self.temp_comp_a) == 0 and int(self.temp_comp_b) == 0:
                    self.temp_comp = "None"
                if int(self.temp_comp_a) == 1 and int(self.temp_comp_b) == 0:
                    self.temp_comp = "Manual"
                if int(self.temp_comp_a) == 2 and int(self.temp_comp_b) == 0:
                    self.temp_comp = "Automatic"
                elif int(self.temp_comp_a) == 3 and int(self.temp_comp_b) == 0:
                    self.temp_comp = "Linear"
                elif int(self.temp_comp_a) == 3 and int(self.temp_comp_b) == 1:
                    self.temp_comp = "Structural"
                else:
                    self.temp_comp = "Unknown"
            self.isomac_exists = True
            return True
        else:
            self.isomac_exists = False
            return False

    # get serv data from last uploaded serv file in autotune/AutotuneDC. Gets the strokes and line 16/17 values.
    def get_serv_data(self):
        try:
            at_dir = os.listdir(self.cal_config.get_autotune_dir())
            for item in at_dir:
                # Find and open serv file.
                if item.lower() == "serv.stp":
                    self.serv_file_exists = True
                    with open(os.path.join(self.cal_config.get_autotune_dir(), item)) as f:
                        for i in range(1, 8):
                            sizes = f.readline()
                        sizes = str(sizes).strip('\n')
                        sizes = sizes.replace('.0', '')
                        sizes = sizes.split(' ')
                        for i in range(1, 9):
                            f.readline()
                        line_sixteen = f.readline()
                        self.serv_line_sixteen = line_sixteen.strip('\n')
                        self.serv_line_sixteen_values = self.serv_line_sixteen.split(" ")[2:4]
                        line_seventeen = f.readline()
                        self.serv_line_seventeen = line_seventeen.strip('\n')
                        self.serv_line_seventeen_values = self.serv_line_seventeen.split(" ")[2:4]
                        self.serv_lxx_stroke = sizes[1]
                        self.serv_lyy_stroke = sizes[2]
                        self.serv_lzz_stroke = sizes[3]
                        print(f'Machine data from uploaded serv file. Location: '
                              f'{self.cal_config.get_autotune_dir()}\nLyy(Bridge axis): {self.serv_lyy_stroke}\n'
                              f'Lxx(Long axis):   {self.serv_lxx_stroke}\nLzz:              {self.serv_lzz_stroke}\n'
                              f'Granite: {self.serv_line_sixteen}\nBridge:  {self.serv_line_seventeen}')
            if not os.path.exists(self.cal_config.get_autotune_dir() + "/serv.stp") and \
                    not os.path.exists(self.cal_config.get_autotune_dir() + "/Serv.stp"):
                # print(f"serv not found in {self.cal_config.get_autotune_dir()}")
                self.serv_file_exists = False
        except FileNotFoundError as e:
            print(e)

    def get_uncertainty(self):
        if os.path.exists("C:/Program Files (x86)/DeaReport/ATReport.ini"):
            with open("C:/Program Files (x86)/DeaReport/ATReport.ini") as f:
                _data = f.readlines()
                _data = _data[2].strip("\n")
                # print(f"Uncertainty: {_data}")
                if _data == "DisabilitaIncertezza=1":
                    self.is_uncertainty = False
                    print("Uncertainty: Disabled")
                else:
                    self.is_uncertainty = True
                    print("Uncertainty: Enabled")
            self.uncertainty_exists = True
        else:
            self.uncertainty_exists = False
            print("Uncertainty data not available")

    # Check uploaded machine params against stroke of serv file
    def run_serv_stroke_test(self):
        x, y, z = 0, 0, 0
        x_true, y_true, z_true = False, False, False
        controller_type = self.cal_config.get_controller_type()
        try:
            # Gets the size of the strokes
            if controller_type == 'DC':
                # Todo modify bravo and vento if statement (IS TIGO OK?)
                if self.cal_config.model == "Bravo":
                    x = eT.parse('C:\\AutotuneDC\\paramX.xml')
                # elif machino == "Vento":
                #     myTree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                else:
                    x = eT.parse('C:\\AutotuneDC\\paramX.xml')
                    y = eT.parse('C:\\AutotuneDC\\paramY.xml')
                    z = eT.parse('C:\\AutotuneDC\\paramZ.xml')
                max_x_stroke = x.findall('maxStrokeSw/Value')
                min_x_stroke = x.findall('minStrokeSw/Value')
                max_y_stroke = y.findall('maxStrokeSw/Value')
                min_y_stroke = y.findall('minStrokeSw/Value')
                max_z_stroke = z.findall('maxStrokeSw/Value')
                min_z_stroke = z.findall('minStrokeSw/Value')
                x = abs(float(min_x_stroke[0].text) - float(max_x_stroke[0].text))
                y = abs(float(min_y_stroke[0].text) - float(max_y_stroke[0].text))
                z = abs(float(min_z_stroke[0].text) - float(max_z_stroke[0].text))
            if controller_type == 'CC' or controller_type == "CC_Hyper":
                with open('C:\\Autotune\\CONSTANT.ASC') as f:
                    constant = f.readlines()
                    if self.cal_config.model == "Vento":
                        y = f.readlines()[40]
                    elif self.cal_config.model == "Bravo":
                        y = f.readlines()[40]
                    else:
                        x = abs(float(constant[40]) - float(constant[33]))
                        y = abs(float(constant[41]) - float(constant[34]))
                        z = abs(float(constant[42]) - float(constant[35]))
                        print(x, y, z)
            print("\nFrom uploaded serv file.", "From constant/params file")
            print(self.serv_lyy_stroke, x)
            print(self.serv_lxx_stroke, y)
            print(self.serv_lzz_stroke, z)
            if (x * 0.9) <= float(self.serv_lyy_stroke) <= (x * 1.1):
                x_true = True
            if (y * 0.9) <= float(self.serv_lxx_stroke) <= (y * 1.1):
                y_true = True
            if (z * 0.9) <= float(self.serv_lzz_stroke) <= (z * 1.1):
                z_true = True
            if x_true and y_true and z_true:
                print("Returned as true. Equality match for serv and last uploaded controller stroke params")
                self.serv_stroke_equality = True
            else:
                print("Returned as false. No equality match for serv and last uploaded controller stroke params")
                self.serv_stroke_equality = False
            return True
        except Exception as e:
            print(f"Error is reported as: {e}")
            return False

    # check serv line 16 and 17 against isomac temp comp data
    def run_serv_atreport_correlation(self):
        if float(self.serv_line_sixteen_values[0]) != 0 or float(self.serv_line_sixteen_values[1]) != 0:
            # print("temp comp detected from line 16 delta values\n")
            if self.temp_comp == "Linear" or self.temp_comp == 'Automatic' or self.temp_comp == 'Manual' \
                    or self.temp_comp == 'None':
                # print(f'self.temp_comp: {self.temp_comp}.')
                return False
            else:
                return True
        else:
            if float(self.serv_line_sixteen_values[0]) == 0 and float(self.serv_line_sixteen_values[1]) == 0:
                if self.temp_comp == "Linear" or self.temp_comp == 'Automatic' or self.temp_comp == 'Manual' \
                        or self.temp_comp == 'None':
                    return True
                else:
                    return False

    def check_if_scanning_probe(self):
        # check if probe is a scanning probe
        scanning = False
        iso10360_4_text_list = []
        if os.path.exists('C:/PCDMISW/PCDDATA/IsoUt.dat'):
            with open('C:/PCDMISW/PCDDATA/IsoUt.dat') as f:
                probe_type = f.readlines()
                self.probe = probe_type[4].strip('\n')
                print(f'self.probe is: {self.probe}')
                scan_probe_list = ['sp25', 'sp600', 'x1', 'hh-p-x1c', 'lsp-x1', 'lsp-x1h', 'x3', 'x5', 'hp-s']
                # put self.probe in a list
                if any(substring in self.probe.lower() for substring in scan_probe_list):
                    scanning = True
                    print(f'p found. scanning is: {scanning} with a {self.probe}')
                else:
                    print('probe not found in scanning list')

        #  get customer data from the ISO103604INFO.TXT file and put into a list iso10360_4_text_list
        if os.path.exists('C:/PCDMISW/iso10360-4/ISO103604INFO.TXT'):
            with open('C:/PCDMISW/iso10360-4/ISO103604INFO.TXT') as f:
                iso10360_4_text = f.readlines()
                for item in iso10360_4_text:
                    iso10360_4_text_list.append(item.strip('\n'))
                # print(f'Model is {self.serial_no} in -4info.txt: {iso10360_4_text_list}')

        # if scanning probe == true, check ISO103604INFO.TXT customer data against serial number/customer and model
        if scanning:
            print(f'if scanning: {self.model}, {self.serial_no}')
            if self.serial_no in iso10360_4_text_list:
                print('---- found')
                return True
            else:
                print('---- not found')
                return False
        else:
            return True

    # check interfac against controller type

    # @staticmethod
    def check_gcomp_map_type(self):
        # todo not finished. May need to address none and linear as they only look at atreoprt and not at the serv file
        # todo possibly incorporate a class variable for sensors. ?put in a list and check if sensor num (0.0) is over 0
        # open gcomp config ini file
        gcomp_ini_file = []
        if os.path.exists('C:/servicedea/gcomp2/deaconf.ini'):
            with open('C:/servicedea/gcomp2/deaconf.ini') as f:
                ini_file = f.readlines()
                for item in ini_file:
                    gcomp_ini_file.append(item.strip('\n'))
                for item in gcomp_ini_file:
                    if item.startswith('THCmpTypeCombo'):
                        gcomp_ini_value = item[-1]
                        # print('-------------------')
                        # print(f'value of temp comp configured in gcomp is: {item[-1]}')
                        self.gcomp_temp_comp = int(gcomp_ini_value)
                        try:
                            if int(gcomp_ini_value) == 0 and self.temp_comp == 'None' and not self.serv_file_exists:
                                return True
                            elif int(gcomp_ini_value) == 1 and self.temp_comp == 'Linear':
                                return True
                            elif int(gcomp_ini_value) == 2 and (float(self.serv_line_sixteen_values[0]) != 0 or
                                                                float(self.serv_line_sixteen_values[1]) != 0) and \
                                    float(self.serv_line_seventeen_values[0]) == 0:
                                return True
                            elif int(gcomp_ini_value) == 3 and float(self.serv_line_seventeen_values[0]) != 0 and \
                                    float(self.serv_line_sixteen_values[0]) == 0:
                                return True
                            elif int(gcomp_ini_value) == 4 and (float(self.serv_line_sixteen_values[0]) != 0 or
                                                                float(self.serv_line_sixteen_values[1]) != 0) and \
                                    len(self.serv_line_seventeen_values) > 1:
                                return True
                            else:
                                return False
                        except Exception as e:
                            print(f'exception for check_gcomp_map_type is {e}')

    # check at report setup of 2009 against uncertainty

    def run_detect_cycle(self):
        self.alert = []
        # diagnostic_report = f"C:/Users/{os.getlogin()}/dataAssessment{dt.now().strftime('_%H.%M.%S_%d-%m-%Y')}.txt"
        # print(f'Hi {os.getlogin().split(".")[0]},')
        self.alert.append(f'Hi {os.getlogin().split(".")[0]},')
        # print(f"Your Diagnostic System Assessment on {dt.now().strftime('%d/%m/%Y')} at "
        #       f"{dt.now().strftime('%H:%M:%S')}.")
        self.alert.append(f"Your Diagnostic System Assessment on {dt.now().strftime('%d/%m/%Y')} at "
                          f"{dt.now().strftime('%H:%M:%S')}.")
        # print("---------------------------\nCurrent calibration configuration machine data:")
        # print("----------------Running Detect Cycle----------------")
        self.alert.append("\n----------------Running Detect Cycle----------------")
        if self.get_isomac():
            # print('An ATReport configuration has been detected:-')
            self.alert.append('An ATReport configuration has been detected:-')
            self.show_diagnostics_class_data()
        else:
            print("IsoMac.dat not found.")
        # print("---")
        self.alert.append('---')
        self.detect_pcdmis_versions()
        # print("---")
        self.alert.append('---')
        self.detect_pcdmis_interfacs()
        # print("---")
        self.alert.append('---')
        self.detect_autotune_version()
        # print("---")
        self.alert.append('---')
        self.get_uncertainty()
        # print("---")
        self.alert.append('---')
        self.get_serv_data()
        print("----------------Detect Cycle Complete----------------")

    def run_analysis(self):
        print("\n-----------------Analysis-----------------")
        if not self.detect_autotune_files_exist():
            print('\n--ALERT-- Unable to locate uploaded machine files, therefore unable to run full diagnostics')
        if not self.isomac_exists:
            print('\n--ALERT-- Configure ATReport if a calibration is required. No ATReport machine data found.')
        if self.serv_file_exists:
            # self.run_serv_stroke_test() returns true if uploaded machine params are found in selected autotune.
            if self.run_serv_stroke_test():
                if not self.serv_stroke_equality:
                    print("\n--ALERT-- Serv File: Last uploaded serv file does not correlate with last uploaded "
                          "machine params. Stroke equality is False.")
                if not self.run_serv_atreport_correlation():
                    print(
                        "\n--ALERT-- ATReport: ATReport is not setup with the correct temp comp type according to last "
                        "uploaded serv file and may result in an incorrect temp comp type on generated certificate")
                if not self.check_gcomp_map_type():
                    print("\n--ALERT-- GComp: Temp comp type configured in gcomp may not match the uploaded serv file. "
                          "Or if temp comp is None a residual serv file may be present. "
                          "Check Gcomp is configured correctly for the uploaded serv file or this can result in an "
                          "incorrect configuration being downloaded to the controller.")
            else:
                print('\n--ALERT-- No uploaded machine files detected, system therefore unable to run full diagnostics')
        else:
            # run self.check_gcomp_map_type() just to get the gcomp_temp_comp value.
            # todo add check if uploaded autotune files exist
            self.check_gcomp_map_type()
            if self.temp_comp == 'Structural':
                print(
                    "\n--ALERT-- ATReport configured for Structural temperature compensation. A Serv file is required.")
            if self.gcomp_temp_comp == 2 or self.gcomp_temp_comp == 3 or self.gcomp_temp_comp == 4:
                print("\n--Alert-- GComp is currently set to structural thermal compensation but no uploaded "
                      "serv file has been detected.")
            if self.temp_comp == 'Linear':
                print("Note - No Serv file detected. ATReport is set to Linear.")
                if self.gcomp_temp_comp != 1:
                    print('--Alert-- GComp is not configured for linear temperature compensation.')
                # if self.gcomp_temp_comp:
            else:
                print('\n--Note-- No Serv file detected.')
        if not self.check_if_scanning_probe() and self.model != 'empty' and self.isomac_exists:
            print("\n--ALERT-- Customer Data in ISO1036004INFO.txt may not match the current customer and may result "
                  "in an incorrect -4 pdf report being generated.")
            self.alert.append("\n--ALERT-- Customer Data in ISO1036004INFO.txt may not match the current customer and "
                              "may result in an incorrect -4 pdf report being generated.")
            os.startfile('C:/PCDMISW/iso10360-4/ISO103604INFO.TXT')
        print("\n------------Analysis Complete-------------")
        self.generate_report_document()

        # save to file
        # f = open(diagnostic_report, "w+")
        # f.write("Hi there")
        # os.startfile(diagnostic_report)

    def generate_report_document(self):
        # open('diagnostic_report', "w+")
        with open(f'C:/Users/lee.maloney/diagnostic_report.txt', 'w+') as f:
            for alert in self.alert:
                f.write(alert + '\n')
        os.startfile('C:/Users/lee.maloney/diagnostic_report.txt')
