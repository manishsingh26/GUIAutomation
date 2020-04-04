
import os
import sys
import psutil
import pyautogui
from win32api import GetKeyState
from win32con import VK_CAPITAL


class CheckOperation(object):

    def __init__(self, app_config):
        self.app_config = app_config
        self.caps_button = "capslock"
        self.kill_ras = "TASKKILL /F /IM RASClient.exe"
        self.dynamic_timing = None
        self.system_width_check = self.app_config["system_criteria"]["resolution"]["width"]
        self.system_height_check = self.app_config["system_criteria"]["resolution"]["height"]

    def resolution_check(self):
        screen_width, screen_length = pyautogui.size().width, pyautogui.size().height
        if screen_width != self.system_width_check and screen_length != self.system_height_check:
            print("Please check the system resolution, system resolution should be 1920 X 1080.")
            sys.exit(0)

    def caps_lock_check(self):
        caps_lock_status = GetKeyState(VK_CAPITAL)
        if caps_lock_status == 1:
            print("Caps Lock is ON, Turning Caps Lock OFF.")
            pyautogui.press([self.caps_button])

    def ras_software_check(self):
        running_tools = os.popen("tasklist /v").read().strip().split("\n")
        for each_app in running_tools:
            if "RASClient.exe" in str(each_app):
                print("RASClient is already running. Closing the current running RASClient Sofftware.")
                os.system(self.kill_ras)

    def ram_check(self):
        system_ram = round(psutil.virtual_memory().total / 1073741824)
        if system_ram == 16:
            self.dynamic_timing = 1
        elif system_ram == 8:
            self.dynamic_timing = 1.5
        elif system_ram == 4:
            self.dynamic_timing = 2
        self.app_config["system_criteria"]["dynamic_time"] = self.dynamic_timing
