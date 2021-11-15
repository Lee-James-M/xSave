import pyautogui
import time
import os
import tkinter.messagebox
import xml.etree.ElementTree as eT  # for autotuneDC xml params file


class MachineAction:

    def __init__(self):
        print('Machine Class created')

    @staticmethod
    def logon_to_telnet(controller):
        time.sleep(0.25)
        if controller == 'DC':
            os.startfile("C:\\Users\\Public\\Desktop\\DesktopFDC\\Controller.lnk")
        if controller == 'CC':
            os.startfile("C:/servicedea/hyperterminal/ShortcutCOM1")
        if controller == "CC_Hyper":
            os.startfile("C:/servicedea/hyperterminal/CC_Hyper")
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'e')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'b')
        time.sleep(0.1)
        pyautogui.hotkey('enter')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'e')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'b')
        time.sleep(0.1)

    @staticmethod
    def home_machine(cal_config):
        home = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to home the machine?")
        if home:
            MachineAction.logon_to_telnet(cal_config.get_controller_type())
            pyautogui.write('autzer', interval=0.05)
            pyautogui.hotkey('enter')
            time.sleep(5.0)
            pyautogui.hotkey('alt', 'f4')
            pyautogui.hotkey('enter')
            print("HOMING")
            time.sleep(1.0)

    @staticmethod
    def move_machine_to_front(cal_config):
        mov_mac = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to move machine to the front?")
        if mov_mac:
            MachineAction.logon_to_telnet(cal_config.get_controller_type())
            if cal_config.model == "Vento":
                pyautogui.write('MOVABS 10,,,,', interval=0.05)
                print('MOVABS 10,,,,')
            elif cal_config.model == "Bravo":
                pyautogui.write('MOVABS 10,,,,', interval=0.05)
                print('MOVABS 10,,,,')
            elif cal_config.model == "Tigo":
                tkinter.messagebox.showinfo(title="Info", message="Machine Type not supported")
            else:
                pyautogui.write('MOVABS ,10,,,', interval=0.05)
                print('MOVABS ,10,,,')
            pyautogui.hotkey('enter')
            time.sleep(6.0)
            pyautogui.hotkey('alt', 'f4')
            pyautogui.hotkey('enter')
            print("CMM TO FRONT")
            time.sleep(1.0)

    @staticmethod
    def move_cmm_to_rear(y_move_to_position, cal_config):
        MachineAction.logon_to_telnet(cal_config.get_controller_type())
        if cal_config.model == "Vento":
            pyautogui.write('MOVABS ' + str(y_move_to_position)[0:8] + ",,,,", interval=0.05)
            print('MOVABS ' + str(y_move_to_position)[0:8] + ",,,,")
        elif cal_config.model == "Bravo":
            pyautogui.write('MOVABS ' + str(y_move_to_position)[0:8] + ",,,,", interval=0.05)
            print('MOVABS ' + str(y_move_to_position)[0:8] + ",,,,")
        elif cal_config.model == "Tigo":
            tkinter.messagebox.showinfo(title="Info", message="Machine Type not supported")
        else:
            pyautogui.write('MOVABS ,' + str(y_move_to_position)[0:8] + ",,,", interval=0.05)
            print('MOVABS ,' + str(y_move_to_position)[0:8] + ",,,")
        pyautogui.hotkey('enter')
        time.sleep(6.0)
        pyautogui.hotkey('alt', 'f4')
        pyautogui.hotkey('enter')
        print("CMM TO REAR")
        time.sleep(1.0)

    @staticmethod
    def move_machine_to_back(cal_config):
        mov_mac = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to move machine to the back?")
        if mov_mac:
            controller_type = cal_config.get_controller_type()
            print("Controller type = " + str(controller_type))
            try:
                if controller_type == 'DC':
                    if cal_config.model == "Bravo":
                        my_tree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                    # elif machino == "Vento":
                    #     my_tree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                    else:
                        my_tree = eT.parse('C:\\AutotuneDC\\paramY.xml')
                    # treeRoot = my_tree.getroot()
                    max_sw_stroke = my_tree.findall('maxStrokeSw/Value')
                    print("sw maxstroke is: " + str(max_sw_stroke[0].text))
                    y_to_move = (float(max_sw_stroke[0].text) * 0.95)
                    print(y_to_move)
                    os.startfile("C:\\Users\\Public\\Desktop\\DesktopFDC\\Controller.lnk")
                if controller_type == 'CC' or controller_type == "CC_Hyper":
                    with open('C:\\Autotune\\CONSTANT.ASC') as f:
                        if cal_config.model == "Vento":
                            cont_y = f.readlines()[40]
                        elif cal_config.model == "Bravo":
                            cont_y = f.readlines()[40]
                        else:
                            cont_y = f.readlines()[41]
                        print("sw maxstroke is: " + cont_y)
                        print(float(cont_y) * 0.95)
                        y_to_move = (float(cont_y) * 0.95)
                        print("y_to_move is: " + str(y_to_move))
                        if controller_type == "CC":
                            os.startfile("C:/servicedea/Hyperterminal/ShortcutCOM1")
                        if controller_type == "CC_Hyper":
                            os.startfile("C:/servicedea/hyperterminal/CC_Hyper")
                MachineAction.move_cmm_to_rear(y_to_move, cal_config)
            except FileNotFoundError as err:
                tkinter.messagebox.showinfo(title="Info", message=f'{err}. Upload autotune files to run this function')
                controller_type = cal_config.get_controller_type()
                print("File not found, Controller type number = " + str(controller_type))
                # TODO Machine move to back - check this second try statement by forcing exception on next cal. ?Add cont type== DC_Hyper?
                try:
                    if controller_type == 'DC':
                        if cal_config.model == "Bravo":
                            my_tree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                        # elif machino == "Vento":
                        #     my_tree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                        else:
                            my_tree = eT.parse('C:\\AutotuneDC\\paramY.xml')
                        # treeRoot = my_tree.getroot()
                        max_sw_stroke = my_tree.findall('maxStrokeSw/Value')
                        print("sw maxstroke is: " + str(max_sw_stroke))
                        y_to_move = (float(max_sw_stroke[0].text) * 0.9)
                        print(y_to_move)
                        os.startfile("C:\\Users\\Public\\Desktop\\DesktopFDC\\Controller.lnk")
                    if controller_type == 'CC':
                        with open('C:\\Autotune\\CONSTANT.ASC') as f:
                            if cal_config.model == "Vento":
                                cont_y = f.readlines()[40]
                            elif cal_config.model == "Bravo":
                                cont_y = f.readlines()[40]
                            else:
                                cont_y = f.readlines()[41]
                            print("sw maxstroke is: " + cont_y)
                            print(float(cont_y) * 0.9)
                            y_to_move = float(cont_y) * 0.9
                            print("y_to_move is: " + str(y_to_move))
                            os.startfile("C:/servicedea/Hyperterminal/ShortcutCOM1")
                    MachineAction.move_cmm_to_rear(y_to_move, cal_config)
                except FileNotFoundError:
                    tkinter.messagebox.showinfo(title="Info", message="Upload autotune files")
                    print("File not found")

    @staticmethod
    def move_machine_to_top(cal_config):
        mov_mac = tkinter.messagebox.askyesno("Machine movement", "Are you sure you want to move Z to the top?")
        if mov_mac:
            MachineAction.logon_to_telnet(cal_config.get_controller_type())
            pyautogui.write('initcm', interval=0.05)
            pyautogui.hotkey('enter')
            print("Reboot of CMM - INITCM")
            time.sleep(1.0)
            pyautogui.write('MOVABS ,,-15,,', interval=0.05)
            print('MOVABS ,,-15,,')
            pyautogui.hotkey('enter')
            time.sleep(5.0)
            pyautogui.hotkey('alt', 'f4')
            pyautogui.hotkey('enter')
            print("CMM TO TOP")
            time.sleep(1.0)
