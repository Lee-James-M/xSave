# reboot etc
import time
import os
import pyautogui
import tkinter.messagebox
import xml.etree.ElementTree as eT  # for autotuneDC xml params file
from selenium import webdriver
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import NoSuchElementException

deps = "C:/Users/lee.maloney/OneDrive - Hexagon/Deps_xSave/"


class ControllerAction:

    def __init__(self):
        print('Controller Class created')

    @staticmethod
    def logon_to_telnet(controller):
        time.sleep(0.2)
        print(f'Controller in telnet log on is {controller}')
        if controller == 'DC':
            os.startfile("C:\\Users\\Public\\Desktop\\DesktopFDC\\Controller.lnk")
        if controller == 'CC':
            os.startfile("C:/servicedea/hyperterminal/ShortcutCOM1")
        if controller == "CC_Hyper":
            os.startfile("C:/servicedea/hyperterminal/CC_Hyper")
        time.sleep(0.25)
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
    def open_fdc():
        driver = webdriver.Chrome(executable_path=os.path.join(deps, 'chromedriver_win32/chromedriver.exe'))
        driver.get("http://100.0.0.1")
        driver.maximize_window()
        # Log-in
        time.sleep(3.0)
        try:
            driver.find_element_by_id("login_PassText").send_keys("hexmet")
            time.sleep(0.5)
        except ElementNotInteractableException:
            driver.find_element_by_id("errorButton1").click()
            time.sleep(0.5)
            driver.find_element_by_id("login_PassText").send_keys("hexmet")
            time.sleep(0.5)
        driver.find_element_by_id("login_BtnLogin").click()
        time.sleep(1.0)
        driver.find_element_by_link_text('Errors History').click()
        time.sleep(1.0)
        try:
            driver.find_element_by_xpath('//*[@id="Button3"]').click()
        except ElementNotInteractableException as e:
            print(f'e1: {e}')
        except NoSuchElementException as e:
            print(f'e2: {e}')

        # except ElementNotInteractableException as e:
        #     print(e)
        #     time.sleep(2.0)
        #     driver.find_element_by_id("errorButton1").click()
        #     time.sleep(1.5)
        #     driver.find_element_by_id("login_PassText").send_keys("hexmet")
        #     time.sleep(0.5)
        #     driver.find_element_by_id("login_BtnLogin").click()
        #     time.sleep(1.5)
        #     driver.find_element_by_link_text('Errors History').click()
        #     time.sleep(1.5)
        #     try:
        #         driver.find_element_by_xpath('//*[@id="Button3"]').click()
        #     except ElementNotInteractableException as e:
        #         print(e)

        # try:
        #     time.sleep(2.5)
        #     driver.find_element_by_id("login_PassText").send_keys("hexmet")
        #     time.sleep(1.5)
        #     driver.find_element_by_id("login_BtnLogin").click()
        #     time.sleep(1.5)
        #     driver.find_element_by_link_text('Errors History').click()
        #     time.sleep(1.5)
        #     try:
        #         driver.find_element_by_xpath('//*[@id="Button3"]').click()
        #     except ElementNotInteractableException as e:
        #         print(e)
        # except ElementNotInteractableException as e:
        #     print(e)
        #     time.sleep(2.0)
        #     driver.find_element_by_id("errorButton1").click()
        #     time.sleep(1.5)
        #     driver.find_element_by_id("login_PassText").send_keys("hexmet")
        #     time.sleep(0.5)
        #     driver.find_element_by_id("login_BtnLogin").click()
        #     time.sleep(1.5)
        #     driver.find_element_by_link_text('Errors History').click()
        #     time.sleep(1.5)
        #     try:
        #         driver.find_element_by_xpath('//*[@id="Button3"]').click()
        #     except ElementNotInteractableException as e:
        #         print(e)

    @staticmethod
    def reboot_common_controller(controller):
        if controller == 'CC' or controller == "CC_Hyper":
            ControllerAction.logon_to_telnet(controller)
            pyautogui.hotkey('ctrl', 'y')
            time.sleep(0.2)
            pyautogui.write('testsoft', interval=0.05)
            pyautogui.hotkey('enter')
            time.sleep(0.5)
            pyautogui.write('bns', interval=0.05)
            pyautogui.hotkey('enter')
            time.sleep(0.2)
            pyautogui.write('9', interval=0.05)
            pyautogui.hotkey('enter')
            time.sleep(0.2)
            pyautogui.write('1', interval=0.05)
            pyautogui.hotkey('enter')
            time.sleep(0.2)
            pyautogui.write('y', interval=0.05)
            pyautogui.hotkey('enter')
            time.sleep(2.0)

    @staticmethod
    def run_initcm_command(controller):
        ControllerAction.logon_to_telnet(controller)
        pyautogui.write('initcm', interval=0.05)
        pyautogui.hotkey('enter')
        print("Reboot of CMM - INITCM")
        time.sleep(1.0)
        pyautogui.hotkey('alt', 'f4')
        pyautogui.hotkey('enter')

    @staticmethod
    def show_cc_controller_params(controller):
        ControllerAction.logon_to_telnet(controller)
        pyautogui.write('show SERIALNO', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)
        pyautogui.write('show CMMCONFG', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)
        pyautogui.write('show ADJFILE', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)
        pyautogui.write('show VMAX', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)
        pyautogui.write('show MAX_ACCL', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)
        pyautogui.write('show M_SWLIM', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)
        pyautogui.write('show P_SWLIM', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)
        pyautogui.write('show HAVE_PH9', interval=0.05)
        pyautogui.hotkey('enter')
        time.sleep(0.5)

    @staticmethod
    def show_machine_status(cal_config):
        ControllerAction.logon_to_telnet(cal_config.get_controller_type())
        pyautogui.write('status', interval=0.05)
        pyautogui.hotkey('enter')

    @staticmethod
    def show_cc_fw_version(cal_config):
        ControllerAction.logon_to_telnet(cal_config.get_controller_type())
        pyautogui.write('version', interval=0.05)
        pyautogui.hotkey('enter')

    @staticmethod
    def move_machine_to_front(cal_config):
        ControllerAction.logon_to_telnet(cal_config.get_controller_type())
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
    def move_cmm_to_rear(y_position_to_move, cal_config):
        ControllerAction.logon_to_telnet(cal_config.get_controller_type())
        if cal_config.model == "Vento":
            pyautogui.write('MOVABS ' + str(y_position_to_move)[0:8] + ",,,,", interval=0.05)
            print('MOVABS ' + str(y_position_to_move)[0:8] + ",,,,")
        elif cal_config.model == "Bravo":
            pyautogui.write('MOVABS ' + str(y_position_to_move)[0:8] + ",,,,", interval=0.05)
            print('MOVABS ' + str(y_position_to_move)[0:8] + ",,,,")
        elif cal_config.model == "Tigo":
            tkinter.messagebox.showinfo(title="Info", message="Machine Type not supported")
        else:
            pyautogui.write('MOVABS ,' + str(y_position_to_move)[0:8] + ",,,", interval=0.05)
            print('MOVABS ,' + str(y_position_to_move)[0:8] + ",,,")
        pyautogui.hotkey('enter')
        time.sleep(6.0)
        pyautogui.hotkey('alt', 'f4')
        pyautogui.hotkey('enter')
        print("CMM TO REAR")
        time.sleep(1.0)

    @staticmethod
    def move_machine_to_back(cal_config):
        controller_type = cal_config.get_controller_type()
        print("Controller type number = " + str(controller_type))
        try:
            if controller_type == 'DC':
                if cal_config.model == "Bravo":
                    my_tree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                # elif machino == "Vento":
                #     myTree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                else:
                    my_tree = eT.parse('C:\\AutotuneDC\\paramY.xml')
                # treeRoot = myTree.getroot()
                max_sw_stroke = my_tree.findall('maxStrokeSw/Value')
                print("sw maxstroke is: " + str(max_sw_stroke))
                y_to_move = (float(max_sw_stroke[0].text) * 0.9)
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
                    print(float(cont_y) * 0.9)
                    y_to_move = float(cont_y) * 0.9
                    print("yToMove is: " + str(y_to_move))
                    os.startfile("C:/servicedea/Hyperterminal/ShortcutCOM1")
            ControllerAction.move_cmm_to_rear(y_to_move, cal_config)
        except FileNotFoundError:
            tkinter.messagebox.showinfo(title="Info", message="Upload autotune files")
            print("File not found")
            controller_type = cal_config.get_controller_type()
            print("Controller type number = " + str(controller_type))
            try:
                if controller_type == 'DC':
                    if cal_config.model == "Bravo":
                        my_tree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                    # elif machino == "Vento":
                    #     myTree = eT.parse('C:\\AutotuneDC\\paramX.xml')
                    else:
                        my_tree = eT.parse('C:\\AutotuneDC\\paramY.xml')
                    # treeRoot = myTree.getroot()
                    max_sw_stroke = my_tree.findall('maxStrokeSw/Value')
                    print("sw maxstroke is: " + str(max_sw_stroke))
                    y_to_move = (float(max_sw_stroke[0].text) * 0.9)
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
                        print(float(cont_y) * 0.9)
                        y_to_move = float(cont_y) * 0.9
                        print("yToMove is: " + str(y_to_move))
                        os.startfile("C:/servicedea/Hyperterminal/ShortcutCOM1")
                ControllerAction.move_cmm_to_rear(y_to_move, cal_config)
            except FileNotFoundError:
                tkinter.messagebox.showinfo(title="Info", message="Upload autotune files")
                print("File not found")

    @staticmethod
    def move_machine_to_top(cal_config):
        ControllerAction.logon_to_telnet(cal_config.get_controller_type())
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
