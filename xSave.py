from tkinter import *
import tkinter.messagebox
from tkinter import ttk
# import os
# import shutil
import datetime
import time  # For sleep function

# from Adafruit_IO import Client
# import win32print  # install pywin32 for this module. Used to change default printer

# import SCSModule as Scs
from CalibrationConfig import *
import MachineActions
import ControllerActions
import ProbeActions
import Datasave
import Email_funcs
import Customer
import Host
import os
# import PyPDF2
# import openpyxl
# from PIL import ImageTk, Image  # for the cal cert
# import xlsxwriter, # import docx, # import gspread, # import docx, # import subprocess
# from openpyxl import load_workbook
# import pyautogui
# import selenium
# for backup email
# from email.mime.base import MIMEBase
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email import encoders
# from selenium import webdriver


def select_controller():
    controller_selection = lb_dc.curselection()
    for item in controller_selection:
        cal_config.controller = lb_dc.get(item)
        lb_controller_type_label.config(text="Controller type is set to: " + cal_config.controller)
        print(lb_controller_type_label.cget("text"))
        lb_controller_type_label.config(fg="green")
        cal_config.put_detail_into_workbook(location='b6', detail=str(cal_config.controller))
    getCurrentConfig()


# Sets listbox label for pcdmis type and saves the selection to the DsName .xlsx.
def select_pcdmis_ver():
    das = lb_pcdmis.curselection()
    for item in das:
        cal_config.set_pcdmis_version(lb_pcdmis.get(item))
        lb_pcdmis_label.config(text="PCDMIS set to: " + cal_config.get_pcdmis_version())
        lb_pcdmis_label.config(fg="green")
        cal_config.put_detail_into_workbook(location='b7', detail=str(cal_config.get_pcdmis_version()))
    getCurrentConfig()
    print(lb_pcdmis_label.cget("text"))


# Sets listbox label for controller type. DC or CC. And saves the selection to the DsName .xlsx.
def select_machine_type():
    das = lb_machinetype.curselection()
    for item in das:
        machine = lb_machinetype.get(item)
        cal_config.model = machine
        lb_machine_label.config(text="Machine type is set to: " + machine)
        lb_machine_label.config(fg="green")
        cal_config.put_detail_into_workbook(location='b5', detail=str(machine))  # from Calibration config - Static
    getCurrentConfig()


# Copies and pastes all backup folders ie autotune, pcddata tutor etc to datasave directory,
# returns autotunedir as AutotuneDC or Autotune.
def copy_all_folders_to_datasave_dir(datasavedir):  # copies files to the datasave directory
    tempcomp_ans = tkinter.messagebox.askyesno("Temp Comp", "Does Machine have Temperature compensation?")
    if tempcomp_ans:
        shutil.copytree("C:\\Program Files\\Thermal_ocx", datasavedir + "Thermal_ocx")
    shutil.copytree("C:\\PCDMISW\\PCDDATA", datasavedir + "PCDDATA")  # copys to new dir
    shutil.copytree("C:\\tutor", datasavedir + "tutor")
    # shutil.copytree(cal_config.get_autotune_dir(), str(datasavedir) + cal_config.autotune_fold_name)
    shutil.copytree(cal_config.get_autotune_dir(), str(datasavedir) + cal_config.get_autotune_c_drive_folder_name())
    shutil.copytree("C:\\servicedea\\cmmdata\\arm1", datasavedir + "cmmdata\\arm1")
    os.mkdir(os.path.join(datasavedir, 'Certificates and Documents'))


# Gets serial and company entries along with machine type to create a datasave folder. Returns datasave directory path.
def gen_datasave_directory_name():
    c = company_entry.get()
    isinstance(c, str)
    backup_name = str(c) + "_" + cal_config.get_model() + "_" + str(date_time) + "\\"
    print(f'backup name is {backup_name}')
    datasavedir = ('C:\\Users\\lee.maloney\\OneDrive - Hexagon\\CalCerts and Health checks\\2021\\'
                   + str(c) + "_" + cal_config.get_model() + "_" + str(date_time) + "\\")
    cal_config.set_datasave_path_name(datasavedir, backup_name)
    return datasavedir


# Runs main element of copy function
def generate_datasave_folder_and_files():
    try:
        datasave_dir = gen_datasave_directory_name()
        datasavedir_exists = os.path.exists(datasave_dir)
        print("Does path already exist: " + str(datasavedir_exists))
        if datasavedir_exists:
            # os.chdir(datasave_dir)
            shutil.rmtree(datasave_dir)
        copy_all_folders_to_datasave_dir(datasave_dir)
        datasave.cleanup_datasave_autotune(datasave_dir + cal_config.get_autotune_c_drive_folder_name())
        print('l1 below')
        print(datasave_dir + cal_config.get_autotune_c_drive_folder_name())
        # cleanup_datasave_autotune(datasave_dir + cal_config.autotune_fold_name)
    except FileNotFoundError:
        print("file not found")
        tkinter.messagebox.showwarning("Error", "Unable to copy files. Ensure you have entered "
                                                "Company_Location_Machine Size_Serial Number data "
                                                "and the correct selections are made in left hand select boxes")


def gen_files_full():
    scan_results_dc()
    time.sleep(0.3)
    cpy_mdb()
    time.sleep(0.3)
    cal_lab_gen()


def set_interfac_to_leitz():
    pcdmis = cal_config.get_pcdmis_directory()
    print(f'pcdmis directory is {pcdmis}')
    path = os.path.exists(pcdmis)
    if path:
        shutil.copyfile(pcdmis + "leitz.dll", pcdmis + "interfac.dll")
        print(f'Interfac set to leitz.dll, PCDMIS path is set to: {pcdmis}')
        notifier_label.config(text="Leitz Interfac Finished")
    else:
        tkinter.messagebox.showerror(title=None, message="Path does not exist for Selected Pcdmis version.")


def set_interfac_to_dc():
    pcdmis = cal_config.get_pcdmis_directory()
    path = os.path.exists(pcdmis)
    if path:
        shutil.copyfile(pcdmis + "FDC.dll", pcdmis + "interfac.dll")
        print(f'Interfac set to FDC.dll, PCDMIS path is set to: {pcdmis}')
        notifier_label.config(text="FDC Interfac Finished")
    else:
        tkinter.messagebox.showerror(title=None, message="Path does not exist for Selected Pcdmis version.")


# Disables uncertainty, uses unCertDisabled string variable pointing to disabled version in Dependencies.
def uncert_dis():
    shutil.copyfile(unCertDisabled, "C:\\Program Files (x86)\\DeaReport\\" + "ATReport.ini")
    print("Uncertainty Disabled")
    notifier_label.config(text="Uncert Disabled Finished")


# Enables uncertainty, uses unCertDisabled string variable pointing to enabled version in Dependencies.
def uncert_en():
    shutil.copyfile(unCertEnabled, "C:\\Program Files (x86)\\DeaReport\\" + "ATReport.ini")
    print("Uncertainty Enabled")
    notifier_label.config(text="Uncert Enabled Finished")


def clear_last_config_and_files():
    delete_all_files = tkinter.messagebox.askyesno("Clear Cal Files", "Are you sure you want to delete all files "
                                                                      "associated with last calibration?")
    if delete_all_files:
        print("Running clear_last_config_and_files")
        dir_list = ["c:\\tutor", "c:/AutotuneDC", "c:/Autotune", "c:\\servicedea\\cmmdata\\arm1", "c:\\Program "
                    "Files\\Thermal_ocx", "c:\\pcdmisw\\PCDDATA", "C:\\DeaReport\\dataexport"]
        for _ in dir_list:
            cal_config.clean_all_working_dirs(_)
        cal_config.clear_workbook_details()
        cal_config.load_workbook_details_to_calconfig_object()
        notifier_label.config(text="Clear files finished")
        lb_controller_type_label.config(text="Controller type not yet selected")  # reset buttons text and color.
        lb_controller_type_label.config(fg="black")
        lb_machine_label.config(text="Machine type not yet selected")
        lb_machine_label.config(fg="black")
        lb_pcdmis_label.config(text="Pcdmis version not yet selected")
        lb_pcdmis_label.config(fg="black")
        notifier_label.config(text="")
        getCurrentConfig()
    else:
        print("Clear last files not run")


# deletes files with specific extensions in datasave autotune directory so only controller files are left
# enabling folder to be emailed without error.
# def cleanup_datasave_autotune(dir_name):
#     test = os.listdir(dir_name)
#     autotune_remove_list = [".exe", ".dll", ".config", ".zip", ".prc", ".ht", ".tab", ".bat"]
#     for item in test:
#         for _ in autotune_remove_list:
#             if item.lower().endswith(_):
#                 os.remove(os.path.join(dir_name, item))
#     for item in test:
#         if item.lower().endswith('data') or item.lower().endswith('v05.5'):
#             shutil.rmtree(os.path.join(dir_name, item))


def insert_recall_name():
    # namerecall = cal_config.datasave_path[71:]
    namerecall = cal_config.get_backup_name()
    fold_name_recall_label.config(text=namerecall)


def open_geotools():
    os.chdir("C:\\Users\\lee.maloney\\OneDrive - Hexagon\\Profile\\Downloads\\Geotools_v13.1")
    os.startfile("C:\\Users\\lee.maloney\\OneDrive - Hexagon\\Profile\\Downloads\\Geotools_v13.1\\"
                 "Geotools_v13.1")


def openfdc():
    controller_action.open_fdc()


def run_cc_autotune():
    os.chdir("C:\\Autotune\\")
    os.startfile("c:\\Autotune\\Autadj.exe")


def run_dc_autotune():
    os.chdir("C:\\AutotuneDC\\")
    os.startfile("c:\\AutotuneDC\\Autadj.exe")


def run_cc_creacos():
    os.chdir("C:\\Autotune\\")
    os.startfile("c:\\Autotune\\Creacos.exe")


def run_dc_creacos():
    os.chdir("C:\\AutotuneDC\\")
    os.startfile("c:\\AutotuneDC\\Creacos.exe")


# checks if controller and machine type are selected and checks if corresponding adj folder is present in Dependencies
# folder. If found will copy ADJ's and params.txt for selected controller autotune directory
def adj_move():
    print("adjmove")
    if cal_config.model is not None:
        print("carry on to dc, m ok")
        if cal_config.get_autotune_dir() is not None:
            print("carry on to s. dc ok")
            s = serial_entry.get()
            cal_config.set_serial_number(s)
            isinstance(s, int)
            if s:
                print("serial entry.get " + s)
                m = cal_config.model.lower()
                mac_list = (
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
                for mac_lower, mac_upper in mac_list:
                    if m == mac_lower:
                        m = mac_upper
                        print(f'm is equal to: {m}')
                adj_fold = deps + "ADJ/" + m + s + "/"
                print("ADJ folder set to " + adj_fold)
                adj = os.path.exists(adj_fold)  # check adj folder with adjs exists
                if adj:
                    item = os.listdir(adj_fold)
                    print(f'items in adj_fold are: {item}, get_autotune_dir() is {cal_config.get_autotune_dir()}')
                    for f in item:
                        shutil.copy(adj_fold + f, cal_config.get_autotune_dir())  # at is autotune directory
                    notifier_label.config(text="Adj Import Finished")
                else:
                    tkinter.messagebox.showwarning("Error",
                                                   "Incorrect machine type match to serial number or serial number "
                                                   "input is not in the right format")
                    tkinter.messagebox.showwarning("Correction",
                                                   "Serial number must be in numerical format with no leading "
                                                   "zeros.")
            else:
                print("no instance of s found.")
                tkinter.messagebox.showwarning("Error", "No serial number input")
        else:
            print("DC is none")
            tkinter.messagebox.showwarning("Error", "Machine controller not selected")
    else:
        print("m is none")
        tkinter.messagebox.showwarning("Error", "Machine type not selected")


def gen_map_code():
    email_func.get_and_email_fdc_code()


# Opens an instance of GComp
def gcomp():
    os.chdir("c:\\servicedea\\gcomp2")
    os.startfile("c:\\servicedea\\gcomp2\\geocomp13.exe")


# Finds serv.stp file based on controller type listbox selection. May need capitals for DC
def cpy_serv():
    cal_config.copy_serv_file()
    notifier_label.config(text="Serv Copy Finished")


def copy_fdc():
    fdcdir = os.listdir(os.path.join(deps, 'FDC-Firmware'))
    print(f'Copyfdc() run. Items listed in fdc dependencies dir are: {fdcdir}')
    fdctarget = "C:\\Users\\Public\\Desktop\\DesktopFDC\\"
    os.startfile(fdctarget + "ftptarget")
    os.startfile(os.path.join(deps, 'FDC-Firmware/'))


def scan_results_dc():
    datasave.move_scan_results_to_datasave(cal_config)
    notifier_label.config(text="Scan Copy Finished")


# Finds mdb and creates a copy
def cpy_mdb():
    datasave.move_mdb_to_datasave(cal_config)
    notifier_label.config(text="MDB Copy Finished")


# Copies a blank cal label and service check sheet.
def cal_lab_gen():
    datasave.generate_cal_label(cal_config)


# send datasave file name to generate function in the imported SCSModule(as Scs).
def scs_gen():
    pass
#     Scs.generate(os.path.join(cal_config.datasave_path, "ServiceCheckSheet.docx"))


def convert_2pdf():
    datasave.convert_files_to_pdf(cal_config)


def zip_datasave():
    datasave.create_datasave_zipfile(cal_config)


def send_zip_to_hex():
    send_zip("ServiceAdmin.uk@hexagon.com")


def send_zip_to_me():
    send_zip("Lee.maloney@hexagon.com")


def send_zip(email_recipient):
    email_func.send_zipped_backup_email(email_recipient)


def open_cc_fw():
    os.startfile(os.path.join(deps, 'REV25_00\\autorun.exe'))


def delete_partprogs():
    # try except PermissionError pcdmis may be using the folder
    delpp = tkinter.messagebox.askyesno("Part program Dir", "Are you sure you want to delete PP directory?")
    if delpp:
        try:
            shutil.rmtree("C:\\pcdmisw\\pcdpp")
            print("PP Dir deleted")
        except FileNotFoundError:
            tkinter.messagebox.showinfo("info", "File already deleted or does not exit")
    else:
        print("PP Dir Not deleted")


def delete_probes():
    # try except PermissionError pcdmis may be using the folder
    pcdmis = cal_config.get_pcdmis_directory()
    probesDir = ""
    if pcdmis == "C:\\Program Files\\Hexagon\\PC-DMIS 2019 R2 64-bit\\":
        probesDir = "C:\\Users\\Public\\Documents\\Hexagon\\PC-DMIS\\2019 R2"
    if pcdmis == "C:\\Program Files\\Hexagon\\PC-DMIS 2020 R1 64-bit\\":
        probesDir = "C:\\Users\\Public\\Documents\\Hexagon\\PC-DMIS\\2020 R1"
    if pcdmis == "C:\\Program Files\\Hexagon\\PC-DMIS 2020 R2 64-bit\\":
        probesDir = "C:\\Users\\Public\\Documents\\Hexagon\\PC-DMIS\\2020 R2"
    print("it is " + str(pcdmis))
    _del_probes = tkinter.messagebox.askyesno("Part PCDMIS Probe Files", "Are you sure you want to delete probe files?")
    if _del_probes:
        probe_path = os.path.exists(probesDir)
        if probe_path:
            test = os.listdir(probesDir)
            for item in test:
                if item.endswith(".PRB") or item.endswith(".Results"):
                    os.remove(os.path.join(probesDir, item))
        else:
            print("Probe path does not exist")


def get_documentation():
    os.startfile(os.path.join(os.path.join(deps, 'Service_aids'),
                              service_aid_combo.get()))
    print(service_aid_combo.get())


def update_aid_combobox():
    doc_list = []
    os.chdir("C:\\Users\\lee.maloney\\OneDrive - Hexagon\\Deps_xSave")
    for root_dir, dirs, files in os.walk(os.path.join(deps, 'Service_aids')):
        for aid_file in files:
            doc_list.append(aid_file)
    service_aid_combo["values"] = doc_list


def open_last_datasave():
    os.startfile(cal_config.get_datasave_path_name())


def getCurrentConfig():
    config_label.config(text=cal_config.model + ", " + cal_config.controller + ", " + cal_config.pcdmis_version)


def rotate_head_90_0():
    rot_head = tkinter.messagebox.askyesno("Rotate Head Command", "Are you sure you want to rotate the head?")
    if rot_head:
        print('user said yes')
        probe_action.rotate_probe_head(90, 0, cal_config.get_controller_type())


def rotate_head_90_180():
    rot_head = tkinter.messagebox.askyesno("Rotate Head Command", "Are you sure you want to rotate the head?")
    if rot_head:
        probe_action.rotate_probe_head(90, 180, cal_config.get_controller_type())


def rotate_head_0_0():
    rot_head = tkinter.messagebox.askyesno("Rotate Head Command", "Are you sure you want to rotate the head?")
    if rot_head:
        probe_action.rotate_probe_head(0, 0, cal_config.get_controller_type())


def probe_head_test_seq():
    rot_head = tkinter.messagebox.askyesno("Rotate Head Command", "Are you sure you want to run probe head test "
                                                                  "sequence?")
    if rot_head:
        probe_action.run_probe_head_test_seq(cal_config.get_controller_type())


def send_machine_front():
    mov_mac = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to move machine to the front?")
    if mov_mac:
        machine_action.move_machine_to_front(cal_config)


def send_machine_back():
    mov_mac = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to move machine to the back?")
    if mov_mac:
        machine_action.move_machine_to_back(cal_config)


def send_machine_top():
    mov_mac = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to move Z to the top?")
    if mov_mac:
        machine_action.move_machine_to_top(cal_config)


def homing_seq():
    home = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to home the machine?")
    if home:
        machine_action.home_machine(cal_config)


def ccContReboot():
    controller_action.reboot_common_controller(cal_config.get_controller_type())


def initCm():
    controller_action.run_initcm_command(cal_config.get_controller_type())


def macStatus():
    controller_action.show_machine_status(cal_config)


def contFwVersion():
    controller_action.show_cc_fw_version(cal_config)


def cc_controller_params():
    controller_action.show_cc_controller_params(cal_config.get_controller_type())


def cc_controller_temps():
    pass
    # shutil.copy("c:\\DeaReport", Dependencies + "ATReport Backup\\DeaReport")
    # atdir = os.listdir("c:\\DeaReport\\")
    # print(atdir)
    # for item in atdir:
    #     shutil.copyfile("c:\\DeaReport\\" + item, Dependencies + "ATReport Backup\\DeaReport\\")
    # d_path = os.path.exists(Dependencies + "ATReport Backup\\DeaReport")
    # if d_path:
    #     shutil.rmtree(Dependencies + "ATReport Backup\\DeaReport")
    # shutil.copy("c:\\DeaReport", Dependencies + "ATReport Backup\\DeaReport")

    # ecbSeq()
    # pyautogui.write('READTP', interval=0.05)
    # pyautogui.hotkey('enter')
    # time.sleep(0.5)


def activateMotors():
    pass
#     # feeds = aio.feeds()
#     motorsOn = aio.feeds('motors-on')
#     # motorsOnData = aio.receive(motorsOn.key)
#     aio.send_data(motorsOn.key, "static")
#     aio.send_data(motorsOn.key, "1")
#     time.sleep(5)
#     aio.send_data(motorsOn.key, "static")
#     print("Motors on command sent")


def configHide():
    frame2.pack_forget()


def configShow():
    frame2.pack(side=LEFT)


def operationsHide():
    opsFrame.pack_forget()


def operationsShow():
    opsFrame.pack(side=LEFT)


def version():
    tkinter.messagebox.showinfo(title="Version", message="Current Version is 0.945")


def testo():
    print(Customer.Customer.get_customer_machine_size())
    print(Customer.Customer.get_customer_machine_model())
    print(os.getlogin())
    c = Customer.Customer.get_customer_address()
    print(c)


cal_config = CalConfig('', '', '', '', '', '', '', '', '', '')
probe_action = ProbeActions.ProbeAction()
machine_action = MachineActions.MachineAction()
controller_action = ControllerActions.ControllerAction()
datasave = Datasave.Datasave()
email_func = Email_funcs.Email()
host = Host.Host()

print(f'Name is: {__name__}')

dt = datetime.datetime.now()
date_time = dt.strftime("%d%m%Y")
date_time_mdb = dt.strftime("%Y%m%d")
deps = f'C:/Users/{os.getlogin()}/OneDrive - Hexagon/Deps_xSave/'
print(deps)
unCertDisabled = os.path.join(deps, "Uncertainty/ATReport_wo_un.ini")
unCertEnabled = os.path.join(deps, "Uncertainty/ATReport.ini")

# importing a Tkinter class, this creates a blank window. Its a constructor class. root is our main tkinter object.
root = Tk()
root.iconbitmap(os.path.join(deps, 'Icons/hexIcon.ico'))
root.title("xSave")

# ------------------------------------------------------------------------
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Show Config", command=configShow)
filemenu.add_command(label="Hide Config", command=configHide)
filemenu.add_separator()
filemenu.add_command(label="Show Operations", command=operationsShow)
filemenu.add_command(label="Hide Operations", command=operationsHide)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)

menubar.add_cascade(label="Main", menu=filemenu)

help_menu = Menu(menubar, tearoff=0)
help_menu.add_command(label="Version", command=version)
menubar.add_cascade(label="About", menu=help_menu)

root.config(menu=menubar)
# ------------------------------------------------------------------------, anchor=E
opsFrame = tkinter.Frame(root)
opsFrame.pack(side=LEFT)

probeOps_label = Label(opsFrame, text="Probe Operations")
probeOps_label.config(height=1, width=26)
probeOps_label.grid(row=1)
zeroZero_btn = Button(opsFrame, text="A0, B0", command=rotate_head_0_0)
zeroZero_btn.config(height=1, width=26)
zeroZero_btn.grid(row=2)
ninetyZero_btn = Button(opsFrame, text="A90, B0", command=rotate_head_90_0)
ninetyZero_btn.config(height=1, width=26)
ninetyZero_btn.grid(row=3)
ninetyOneEighty_btn = Button(opsFrame, text="A90, B180", command=rotate_head_90_180)
ninetyOneEighty_btn.config(height=1, width=26)
ninetyOneEighty_btn.grid(row=4)
probeHeadTestSeq_btn = Button(opsFrame, text="Probe Head Test Sequence", command=probe_head_test_seq)
probeHeadTestSeq_btn.config(height=1, width=26)
probeHeadTestSeq_btn.grid(row=5)
Ops1_label = Label(opsFrame, text=" ")
Ops1_label.config(height=1, width=26)
Ops1_label.grid(row=6)

macOps_label = Label(opsFrame, text="Machine Operations")
macOps_label.config(height=1, width=26)
macOps_label.grid(row=7)
motorButton = Button(opsFrame, text="Motors On", fg="red", command=activateMotors)
motorButton.config(height=1, width=26)
motorButton.grid(row=8)
home_btn = Button(opsFrame, text="Home Machine", command=homing_seq)
home_btn.config(height=1, width=26)
home_btn.grid(row=9)
macTop_btn = Button(opsFrame, text="Send the CMM to the Top", command=send_machine_top)
macTop_btn.config(height=1, width=26)
macTop_btn.grid(row=10)
macFront_btn = Button(opsFrame, text="Send the CMM to the Front", command=send_machine_front)
macFront_btn.config(height=1, width=26)
macFront_btn.grid(row=11)
macBack_btn = Button(opsFrame, text="Send the Machine to the Back", command=send_machine_back)
macBack_btn.config(height=1, width=26)
macBack_btn.grid(row=12)
Ops2_label = Label(opsFrame, text=" ")
Ops2_label.config(height=1, width=26)
Ops2_label.grid(row=13)

contOps_label = Label(opsFrame, text="Common Controller Operations")
contOps_label.config(height=1, width=26)
contOps_label.grid(row=14)
contOpsResetBtn = Button(opsFrame, text="Reboot Controller", command=ccContReboot)
contOpsResetBtn.config(height=1, width=26)
contOpsResetBtn.grid(row=15)
initCm_btn = Button(opsFrame, text="initCm", command=initCm)
initCm_btn.config(height=1, width=26)
initCm_btn.grid(row=16)
contOpsFwBtn = Button(opsFrame, text="Get Firmware Version", command=contFwVersion)
contOpsFwBtn.config(height=1, width=26)
contOpsFwBtn.grid(row=17)
contOpsStatusBtn = Button(opsFrame, text="Get Machine Status", command=macStatus)
contOpsStatusBtn.config(height=1, width=26)
contOpsStatusBtn.grid(row=18)
contOpsParamsBtn = Button(opsFrame, text="Main Controller Params", command=cc_controller_params)
contOpsParamsBtn.config(height=1, width=26)
contOpsParamsBtn.grid(row=19)
contOpsTempsBtn = Button(opsFrame, text="Show Controller Temperatures", command=cc_controller_temps)
contOpsTempsBtn.config(height=1, width=26)
contOpsTempsBtn.grid(row=20)
Ops3_label = Label(opsFrame, text=" ")
Ops3_label.config(height=3, width=26)
Ops3_label.grid(row=21)

frame1 = tkinter.Frame(root)
frame1.pack(side=RIGHT)

clear_label = Label(frame1, text="Press to clear last calibration files")
clear_label.grid(row=1, sticky=E)
interfac_label = Label(frame1, text="Select Interfac")
interfac_label.grid(row=2, sticky=E)
uncert_label = Label(frame1, text="Enable uncertainty only for 2009")
uncert_label.grid(row=3, sticky=E)
serial_label = Label(frame1, text="Enter machine S.N. If GLD0245.IA enter only 245.")
serial_label.grid(row=6, sticky=E)
get_map_label = Label(frame1, text="Get map access code from FDC(Experimental)")
get_map_label.grid(row=8, sticky=E)
script2_label = Label(frame1, text="Upload the map and serv file")
script2_label.grid(row=9, sticky=E)
serv_label = Label(frame1, text="Copy uploaded serv.stp to thermal_ocx as serv1.stp")
serv_label.grid(row=10, sticky=E)
fdc_fw_label = Label(frame1, text="Open lastest copy of fdc firmware and ftptarget")
fdc_fw_label.grid(row=11, sticky=E)
creacos_label = Label(frame1, text="Press to run CreaCos Program")
creacos_label.grid(row=12, sticky=E)
autotune_label = Label(frame1, text="Press to run Autotune")
autotune_label.grid(row=13, sticky=E)
cal_machine_label = Label(frame1, text="Calibrate the machine", fg="green")
cal_machine_label.grid(row=15, sticky=E)
company_label = Label(frame1, text="Enter: Company_Location_Machine Size_S.N")
company_label.grid(row=16, sticky=E)
script1_label = Label(frame1, text="Print ATReport50.pdf and ATReport Summary")
script1_label.grid(row=17, sticky=E)
scan_label = Label(frame1, text="Press to copy scanning results to datasave")
scan_label.grid(row=18, sticky=E)
mdb_label = Label(frame1, text="Press to copy database mdb file to datasave")
mdb_label.grid(row=19, sticky=E)
cal_label = Label(frame1, text="Press to generate a cal label")
cal_label.grid(row=20, sticky=E)
conv_label = Label(frame1, text="Convert cal label and service check sheet to PDF")
conv_label.grid(row=21, sticky=E)
zip_label = Label(frame1, text="Create a Zip file")
zip_label.grid(row=22, sticky=E)
zip_label = Label(frame1, text="Send to ServiceAdminUK")
zip_label.grid(row=23, sticky=E)
spacer_label2 = Label(frame1, text="     ")
spacer_label2.grid(row=0, column=3)
config_label = Label(frame1, text="", fg="green")
config_label.grid(row=25, sticky=E)

clr_button = Button(frame1, text="Clear Files and Config      ", fg="red", command=clear_last_config_and_files)
clr_button.config(height=1, width=22)
clr_button.grid(row=1, column=1)
leitz_intfac_button = Button(frame1, text="CC - Set Interfac to Leitz ", command=set_interfac_to_leitz)
leitz_intfac_button.config(height=1, width=22)
leitz_intfac_button.grid(row=2, column=1)
dc_intfac_button = Button(frame1, text="DC - Set Interfac to FDC  ", command=set_interfac_to_dc)
dc_intfac_button.config(height=1, width=22)
dc_intfac_button.grid(row=2, column=2)
uncert_en_button = Button(frame1, text="Enable Uncertainty          ", command=uncert_en)
uncert_en_button.config(height=1, width=22)
uncert_en_button.grid(row=3, column=1)
uncert_dis_button = Button(frame1, text="Disable Uncertainty         ", command=uncert_dis)
uncert_dis_button.config(height=1, width=22)
uncert_dis_button.grid(row=3, column=2)
serial_entry = Entry(frame1)  # create an entry that requires manual input of serial number
serial_entry.grid(row=6, column=1)
adj_button = Button(frame1, text="Import Adj files               ", fg="black", command=adj_move)
adj_button.config(height=1, width=22)
adj_button.grid(row=6, column=2)
gen_map_code_button = Button(frame1, text="Generate map code        ", command=gen_map_code)
gen_map_code_button.config(height=1, width=22)
gen_map_code_button.grid(row=8, column=1)
gcomp_button = Button(frame1, text="Open GComp                  ", command=gcomp)
gcomp_button.config(height=1, width=22)
gcomp_button.grid(row=9, column=1)
serv_button = Button(frame1, text="Copy Serv File                  ", command=cpy_serv)
serv_button.config(height=1, width=22)
serv_button.grid(row=10, column=1)
fdc_fw_button = Button(frame1, text="Copy FDC                         ", command=copy_fdc)
fdc_fw_button.config(height=1, width=22)
fdc_fw_button.grid(row=11, column=1)
run_cc_creacos_btn = Button(frame1, text="Run CC Creacos               ", fg="purple", command=run_cc_creacos)
run_cc_creacos_btn.config(height=1, width=22)
run_cc_creacos_btn.grid(row=12, column=1)
run_dc_creacos_btn = Button(frame1, text="Run DC Creacos                ", fg="purple", command=run_dc_creacos)
run_dc_creacos_btn.config(height=1, width=22)
run_dc_creacos_btn.grid(row=12, column=2)
run_cc_autotune_btn = Button(frame1, text="Run CC Autotune             ", fg="purple", command=run_cc_autotune)
run_cc_autotune_btn.config(height=1, width=22)
run_cc_autotune_btn.grid(row=13, column=1)
run_dc_autotune_btn = Button(frame1, text="Run DC Autotune              ", fg="purple", command=run_dc_autotune)
run_dc_autotune_btn.config(height=1, width=22)
run_dc_autotune_btn.grid(row=13, column=2)
company_entry = Entry(frame1)  # Enter machine_location_size
company_entry.grid(row=16, column=1)

fold_name_recall_btn = Button(frame1, text="Recall Last Name              ", command=insert_recall_name)
fold_name_recall_btn.config(height=1, width=22)
fold_name_recall_btn.grid(row=15, column=2)

fold_name_recall_label = Label(frame1, text="")
fold_name_recall_label.grid(row=15, column=4)

folder_button = Button(frame1, text="Create Datasave Folder    ", command=generate_datasave_folder_and_files)
folder_button.config(height=1, width=22)
folder_button.grid(row=16, column=2)
folder_button2 = Button(frame1, text="Scan, mdb and cal cert    ", command=gen_files_full)
folder_button2.config(height=1, width=22)
folder_button2.grid(row=20, column=2)
scan_button = Button(frame1, text="Copy Scan Results            ", command=scan_results_dc)
scan_button.config(height=1, width=22)
scan_button.grid(row=18, column=1)
mdb_button = Button(frame1, text="Copy mdb Database File ", command=cpy_mdb)
mdb_button.config(height=1, width=22)
mdb_button.grid(row=19, column=1)
cal_label_button = Button(frame1, text="Generate Cal Label          ", command=cal_lab_gen)
cal_label_button.config(height=1, width=22)
cal_label_button.grid(row=20, column=1)
conv_pdf_btn = Button(frame1, text="Convert Files to PDF        ", command=convert_2pdf)
conv_pdf_btn.config(height=1, width=22)
conv_pdf_btn.grid(row=21, column=1)
zipsend_button = Button(frame1, text="Create Zip File                 ", command=zip_datasave)
zipsend_button.config(height=1, width=22)
zipsend_button.grid(row=22, column=1)
send_to_me_btn = Button(frame1, text="Send Backup to Me         ", command=send_zip_to_me)
send_to_me_btn.config(height=1, width=22)
send_to_me_btn.grid(row=23, column=2)
send_to_hex_btn = Button(frame1, text="Send Backup to Hex       ", command=send_zip_to_hex)
send_to_hex_btn.config(height=1, width=22)
send_to_hex_btn.grid(row=23, column=1)
notifier_label = Label(frame1, text="Notifier")
notifier_label.grid(row=24, column=1)

testo_btn = Button(frame1, text="Testo.", command=testo)
testo_btn.config(height=1, width=22)
testo_btn.grid(row=25, column=1)

# Right hand side functions
fdc_btn = Button(frame1, text="Open FDC Panel - I.P = 100.0.0.1", fg="purple", command=openfdc)
fdc_btn.config(height=1, width=28)
fdc_btn.grid(row=1, column=4)
openGeoTools_btn = Button(frame1, text="Open Geo-Tools for Part Programs", fg="purple", command=open_geotools)
openGeoTools_btn.config(height=1, width=28)
openGeoTools_btn.grid(row=3, column=4)
haspSerDong_btn = Button(frame1, text="Open Geo-Tools for Hasp Drivers", fg="purple", command=open_geotools)
haspSerDong_btn.config(height=1, width=28)
haspSerDong_btn.grid(row=6, column=4)
open_cc_fw_btn = Button(frame1, text="Open CC Firmware install program", fg="purple", command=open_cc_fw)
open_cc_fw_btn.config(height=1, width=28)
open_cc_fw_btn.grid(row=9, column=4)
del_pp_btn = Button(frame1, text="Delete part program directory", fg="red", command=delete_partprogs)
del_pp_btn.config(height=1, width=28)
del_pp_btn.grid(row=10, column=4)
del_probes_btn = Button(frame1, text="Delete probe files", fg="red", command=delete_probes)
del_probes_btn.config(height=1, width=28)
del_probes_btn.grid(row=11, column=4)
open_last_datasave_btn = Button(frame1, text="Open last Datasave folder", command=open_last_datasave)
open_last_datasave_btn.config(height=1, width=28)
open_last_datasave_btn.grid(row=16, column=4)

service_aid_combo = ttk.Combobox(frame1, values=["test",
                                                 "Belt tensions DB425IL_E_Rev.17.pdf",
                                                 "belt_tension_pioneer.pdf"],
                                 state="readonly", postcommand=update_aid_combobox)
service_aid_combo.config(width=28)
service_aid_combo.grid(row=18, column=4)
service_aid_btn = Button(frame1, text="Get selected documentation", command=get_documentation)
service_aid_btn.config(height=1, width=28)
service_aid_btn.grid(row=19, column=4)
scs_btn = Button(frame1, text="Generate Service Check Sheet", command=scs_gen)
scs_btn.config(height=1, width=28)
scs_btn.grid(row=20, column=4)

# ADAFRUIT_IO_KEY = "aio_dIrE34HA1VNhlHSr31QEFQw51nem"
# ADAFRUIT_IO_USERNAME = 'LeeJm'
# aio = Client(ADAFRUIT_IO_USERNAME, ADAFRUIT_IO_KEY)

# Create a second frame to the left to include listbox elements---------------------------------------------------------
frame2 = tkinter.Frame(root)
frame2.pack(side=LEFT)

lb_machinetype = Listbox(frame2, width=26, height=15, selectmode=SINGLE)  # create a machine type list box
lb_machinetype.insert(1, "Global AA")
lb_machinetype.insert(2, "Global A")
lb_machinetype.insert(3, "Global B")
lb_machinetype.insert(4, "Global C")
lb_machinetype.insert(5, "Global D")
lb_machinetype.insert(6, "Global E")
lb_machinetype.insert(7, "Global F")
lb_machinetype.insert(8, "Micra")
lb_machinetype.insert(9, "Tigo")
lb_machinetype.insert(10, "Pioneer")
lb_machinetype.insert(11, "Alpha")
lb_machinetype.insert(12, "Delta")
lb_machinetype.insert(13, "Xcel")
lb_machinetype.insert(14, "Chameleon")
lb_machinetype.insert(15, "Vento")
lb_machinetype.insert(16, "Oxalis")
lb_machinetype.insert(17, "Mistral")
lb_machinetype.insert(18, "Bravo")
lb_machinetype.insert(19, "LK")
lb_machinetype.insert(20, "Gamma")
lb_machinetype.insert(21, "Sirrocco")
lb_machinetype.pack()
lb_machine_label = Label(frame2, text="Machine type not yet selected")
lb_machine_label.pack()
lb_machine_button = Button(frame2, text="Press to select machine type", command=select_machine_type)
lb_machine_button.config(height=1, width=26)
lb_machine_button.pack()
space1_label = Label(frame2, text=" ").pack()

lb_dc = Listbox(frame2, width=30, height=2, selectmode=SINGLE)  # create a controller listbox
lb_dc.insert(1, "DC Controller")
lb_dc.insert(2, "CC Controller")
lb_dc.pack()
lb_controller_type_label = Label(frame2, text="Controller type not yet selected")
lb_controller_type_label.pack()
lb_dc_button = Button(frame2, text="Press to select controller type", command=select_controller)
lb_dc_button.config(height=1, width=26)
lb_dc_button.pack()
space2_label = Label(frame2, text=" ").pack()

lb_pcdmis = Listbox(frame2, width=26, height=3, selectmode=SINGLE)  # create  pcdmis list box
lb_pcdmis.insert(1, "PCDMIS 2020 R2")
lb_pcdmis.insert(2, "PCDMIS 2020 R1")
lb_pcdmis.insert(3, "PCDMIS 2019 R2")
lb_pcdmis.insert(4, "PCDMIS 2019 R1 ")
lb_pcdmis.insert(5, "PCDMIS 2018 R2 ")
lb_pcdmis.pack()
lb_pcdmis_label = Label(frame2, text="Pcdmis version not yet selected")
lb_pcdmis_label.pack()
lb_pcdmis_button = Button(frame2, text="Press to select PCDMIS version", command=select_pcdmis_ver)
lb_pcdmis_button.config(height=1, width=26)
lb_pcdmis_button.pack()
space3_label = Label(frame2, text=" ").pack()

config_Button = Button(frame2, text="Show Current Config", command=getCurrentConfig)
config_Button.config(height=1, width=26, padx=25)
config_Button.pack()

# enter code here for name----------------------------------
check_user_legal = host.check_user_legal()
if not check_user_legal:
    root.destroy()
# -----------------------------------------------------------

print("script end")
root.mainloop()  # This will keep the gui running continuously until user closes.
