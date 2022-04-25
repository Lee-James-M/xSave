import time
import os
import pyautogui
from tkinter import simpledialog
import tkinter.messagebox


class ProbeHead:

    def __init__(self):
        print('Probe Class Created')

    @classmethod
    def logon_to_telnet(cls, controller):
        time.sleep(1)
        print(f'controller is set to:{controller}')
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

    @classmethod
    def rotate_head(cls, a_angle, b_angle, controller):
        rot_head = tkinter.messagebox.askyesno("Rotate Head Command", f"Are you sure you want to rotate the head to "
                                                                      f"A{a_angle}, B{b_angle}?")
        if rot_head:
            print('probe head action recieved')
            ProbeHead.logon_to_telnet(controller)
            pyautogui.write('prbhty ph9,' + str(a_angle) + ',' + str(b_angle), interval=0.05)
            pyautogui.hotkey('enter')
            time.sleep(5.0)
            pyautogui.hotkey('alt', 'f4')
            pyautogui.hotkey('enter')

    @classmethod
    def run_test_seq(cls, controller):
        rot_head = tkinter.messagebox.askyesno("Rotate Head Command", "Are you sure you want to run probe head test "
                                                                      "sequence?\nAngles of:\nA0B0,\nA30+-B30,"
                                                                      "\nA30+-B150, "
                                                                      "\nA90+-B90,\nA90B0,\nA90,B180.")
        if rot_head:
            angles = (
                ("0", "0"),
                ("30", "30"),
                ("30", "150"),
                ("30", "-150"),
                ("30", "-30"),
                ("90", "-90"),
                ("90", "90"),
                ("90", "180"),
                ("90", "0"),
            )
            iteration_number = simpledialog.askinteger(title="Number of iterations", prompt="Enter number:")
            print(f"number of iterations: {iteration_number}")
            if iteration_number:
                ProbeHead.logon_to_telnet(controller)
                for _ in range(iteration_number):
                    for aAngle, bAngle in angles:
                        pyautogui.write('prbhty ph9,' + aAngle + ',' + bAngle, interval=0.05)
                        pyautogui.hotkey('enter')
                        time.sleep(7.5)
                pyautogui.hotkey('alt', 'f4')
                pyautogui.hotkey('enter')
