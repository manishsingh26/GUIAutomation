
import sys
import time
import pyautogui
from pywinauto import Application
from win32com.client import Dispatch

from Script.utils.locate_image_check import ImageOperation
from Script.utils.shoot_mail import StatusMail


class AutomationActivity(object):

    def __init__(self, app_config):
        self.app_config = app_config
        self.root_path = self.app_config["root_path"]

        self.speak = Dispatch("SAPI.SpVoice")
        self.img_obj = ImageOperation()
        email_config = self.root_path + "Config/email_config.json"
        self.mail_obj = StatusMail(email_config)

        self.ras_software_path = self.app_config["software_path"]["ras_client"]
        self.ras_user_name = self.app_config["ras_configuration"]["ras_client_credential"]["user_name"]
        self.ras_password = self.app_config["ras_configuration"]["ras_client_credential"]["password"]
        # self.dynamic_timing = float(self.app_config["system_criteria"]["dynamic_time"])
        self.dynamic_timing = 1
        self.logging_dir = self.app_config["logging_path"]
        self.ras_customer = self.app_config["ras_configuration"]["customer_details"]["customer_name"]
        self.ras_ip = self.app_config["ras_configuration"]["customer_details"]["ip"]
        self.ras_connection_reason = self.app_config["ras_configuration"]["customer_details"]["connection_reason"]
        self.ras_remote_user_name = self.app_config["ras_configuration"]["remote_connection"]["user"]
        self.ras_remote_password = self.app_config["ras_configuration"]["remote_connection"]["password"]
        self.ras_net_act_link = self.app_config["ras_configuration"]["net_act"]["link"]
        self.ras_net_act_user_name = self.app_config["ras_configuration"]["net_act"]["user_name"]
        self.ras_net_act_password = self.app_config["ras_configuration"]["net_act"]["password"]
        self.report_directory = self.app_config["report_save_directory"]

        self.tool_opened = self.root_path + self.app_config["image_validation"]["login"]["tool_opened"]
        self.login_image = self.root_path + self.app_config["image_validation"]["login"]["login_check"]
        self.customer_type_search = self.root_path + self.app_config["image_validation"]["customer_search"]["type_area"]
        self.select_button = self.root_path + self.app_config["image_validation"]["customer_search"]["select_button"]
        self.connection_button = self.root_path + self.app_config["image_validation"]["customer_search"]["connection_button"]
        self.nat_address = self.root_path + self.app_config["image_validation"]["customer_search"]["nat_address"]
        self.drop_down_button = self.root_path + self.app_config["image_validation"]["customer_search"]["drop_down_button"]
        self.rdp_select = self.root_path + self.app_config["image_validation"]["customer_search"]["select_rdp"]
        self.connect_button = self.root_path + self.app_config["image_validation"]["customer_search"]["connect_button"]
        self.reason_area = self.root_path + self.app_config["image_validation"]["customer_search"]["reason_area"]
        self.continue_button = self.root_path + self.app_config["image_validation"]["customer_search"]["continue_button"]
        self.remote_desktop = self.root_path + self.app_config["image_validation"]["customer_search"]["remote_app"]
        self.more_choice = self.root_path + self.app_config["image_validation"]["remote_desktop"]["more_choice"]
        self.different_user = self.root_path + self.app_config["image_validation"]["remote_desktop"]["different_user"]
        self.remote_screen1 = self.root_path + self.app_config["image_validation"]["remote_desktop"]["remote_Screen_1"]
        self.remote_screen2 = self.root_path + self.app_config["image_validation"]["remote_desktop"]["remote_screen_2"]
        self.nat_link_opened = self.root_path + self.app_config["image_validation"]["remote_server"]["nat_link_opened"]
        self.report_folder = self.root_path + self.app_config["image_validation"]["remote_server"]["report_folder"]
        self.performance_manager = self.root_path + self.app_config["image_validation"]["remote_server"]["performance_manager"]
        self.navigate = self.root_path + self.app_config["image_validation"]["remote_server"]["navigate"]
        self.saved_reports = self.root_path + self.app_config["image_validation"]["remote_server"]["saved_reports"]
        self.nat_act_connected = self.root_path + self.app_config["image_validation"]["remote_server"]["nat_act_connected"]
        self.export_button = self.root_path + self.app_config["image_validation"]["report_download"]["export_button"]
        self.ready_report = self.root_path + self.app_config["image_validation"]["report_download"]["report_ready"]
        self.download_button = self.root_path + self.app_config["image_validation"]["report_download"]["download_button"]
        self.directory_error = self.root_path + self.app_config["image_validation"]["report_download"]["directory_error"]
        self.download_complete = self.root_path + self.app_config["image_validation"]["report_download"]["download_complete"]
        self.desktop_directory = self.root_path + self.app_config["image_validation"]["report_download"]["desktop_directory"]
        self.locate_file = self.root_path + self.app_config["image_validation"]["report_download"]["file_location"]

    def ras_software_login(self):
        app = Application(backend="win32").start(self.ras_software_path)
        self.app_config["ras_app_trigger"] = app
        time.sleep(10 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "1_ras_client_tool_open.png")
        self.speak.Speak("RAS Client Software Opened.")

        self.img_obj.loop_retry_check(self.tool_opened)

        pyautogui.write(self.ras_user_name, interval=0.02 * self.dynamic_timing)      # RASClient Tool User Name Type
        time.sleep(2 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "2_ras_client_tool_user_name_type.png")

        pyautogui.press("tab")      # Switch to Password Input Box
        time.sleep(2 * self.dynamic_timing)

        pyautogui.write(self.ras_password, interval=0.02 * self.dynamic_timing)       # RASClient Tool User Password Type
        time.sleep(2 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "3_ras_client_tool_password_type.png")
        # self.speak.Speak("RAS Credential Entered.")

        pyautogui.press("enter")        # RASClient Tool Login
        time.sleep(30 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "4_ras_client_tool_login_done.png")

        pyautogui.keyDown("alt")
        pyautogui.keyDown("space")
        pyautogui.press("x")
        pyautogui.keyUp("alt")
        pyautogui.keyUp("space")
        time.sleep(4 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "5_ras_client_tool_maximization.png")
        # self.speak.Speak("RAS Tool Screen Maximized.")

        login_check_1, login_check_2 = self.img_obj.loop_retry_check(self.login_image)
        if login_check_1 is not None and login_check_2 is not None:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Successful")
            return True

        else:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Failed")
            sys.exit(0)

    def customer_config(self):
        customer_type_check_1, customer_type_check_2 = self.img_obj.loop_retry_check(self.customer_type_search)
        x1, y1 = pyautogui.center(customer_type_check_1)
        y1 = y1 + 25
        pyautogui.click(x1, y1, pause=2)      # Customer Name Type Area Click
        pyautogui.screenshot(self.logging_dir + "6_ras_client_tool_customer_search_click.png")

        pyautogui.write(self.ras_customer, interval=0.02 * self.dynamic_timing)  # Customer Name Typing
        pyautogui.screenshot(self.logging_dir + "7_ras_client_tool_customer_name_type.png")
        time.sleep(1 * self.dynamic_timing)

        select_button_check_1, select_button_check_2 = self.img_obj.loop_retry_check(self.select_button)
        pyautogui.click(select_button_check_1)         # Select Button Click
        # self.speak.Speak("Clicking On Select Button.")
        pyautogui.screenshot(self.logging_dir + "8_ras_client_tool_customer_selection_click_position.png")
        time.sleep(15 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "9_ras_client_tool_customer_selection_next_window.png")

        connection_button_check_1, connection_button_button_check_2 = self.img_obj.loop_retry_check(self.connection_button)
        x2, y2 = pyautogui.center(connection_button_check_1)
        y2 = y2 + 20
        pyautogui.click(x2, y2)  # Connection Click
        # self.speak.Speak("Clicking On Connection Button.")
        pyautogui.screenshot(self.logging_dir + "10_ras_client_tool_connection_click.png")
        time.sleep(2 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "11_ras_client_tool_connection_nat_address_search_click.png")

        nat_address_check_1, nat_address_button_check_2 = self.img_obj.loop_retry_check(self.nat_address)
        x3, y3 = pyautogui.center(nat_address_check_1)
        y3 = y3 + 20
        pyautogui.click(x3, y3)     # Click on Search Area NAT Address
        # self.speak.Speak("Clicking On NAT Address Area.")
        time.sleep(3 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "12_ras_client_tool_connection_nat_address_type.png")

        pyautogui.write(self.ras_ip, interval=0.02 * self.dynamic_timing)  # Type NAT Address
        time.sleep(2 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "12_ras_client_tool_connection_nat_address_type.png")

        # drop_down_check_1, drop_down_check_2 = self.img_obj.loop_retry_check(self.drop_down_button)
        # x4, y4 = pyautogui.center(drop_down_check_1)
        # pyautogui.click(x4, y4)      # Application Drop Down
        pyautogui.click(1720, 330)
        pyautogui.screenshot(self.logging_dir + "13_ras_client_tool_connection_application_drop_down.png")
        time.sleep(4 * self.dynamic_timing)

        # rdp_select_check_1, rdp_select_check_2 = self.img_obj.loop_retry_check(self.rdp_select)
        # x5, y5 = pyautogui.center(rdp_select_check_1)
        # pyautogui.click(x5, y5)     # Application RDP Selection
        pyautogui.click(1476, 366)
        pyautogui.screenshot(self.logging_dir + "14_ras_client_tool_connection_application_rdp_selection.png")
        time.sleep(4 * self.dynamic_timing)

        connect_button_check_1, connect_button_check_2 = self.img_obj.loop_retry_check(self.connect_button)
        pyautogui.click(connect_button_check_1)     # Connect Click
        pyautogui.screenshot(self.logging_dir + "15_ras_client_tool_connection_connect_click.png")
        time.sleep(4 * self.dynamic_timing)

        connect_reason_check_1, connect_reason_check_2 = self.img_obj.loop_retry_check(self.reason_area)
        pyautogui.click(connect_reason_check_1)         # Connection Reason Area Click
        pyautogui.screenshot(self.logging_dir + "16_ras_client_tool_connection_reason_click.png")
        time.sleep(4 * self.dynamic_timing)

        pyautogui.write(self.ras_connection_reason, interval=0.02 * self.dynamic_timing)  # Connection Reason Typeing
        time.sleep(4 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "17_ras_client_tool_connection_reason_type.png")

        continue_button_check_1, continue_button_check_2 = self.img_obj.loop_retry_check(self.continue_button)
        pyautogui.click(continue_button_check_1)  # Connection Reason Area Click
        pyautogui.screenshot(self.logging_dir + "18_ras_client_tool_connection_reason_connect_click_position.png")
        time.sleep(30 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "19_ras_client_tool_connection.png")

        connection_check_1, connection_check_2 = self.img_obj.loop_retry_check(self.remote_desktop)
        if connection_check_1 is not None and connection_check_2 is not None:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Successful")
            return True

        else:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Failed")
            sys.exit(0)

    def remote_connection(self):
        more_choice_check_1, more_choice_check_2 = self.img_obj.image_check(self.more_choice)
        pyautogui.click(more_choice_check_1)     # Connect Click
        pyautogui.screenshot(self.logging_dir + "20_ras_client_remote_desktop_more_choices.png")
        time.sleep(4 * self.dynamic_timing)

        different_user_check_1, different_user_check_2 = self.img_obj.image_check(self.different_user)
        pyautogui.click(different_user_check_1)         # Connection Reason Area Click
        pyautogui.screenshot(self.logging_dir + "21_ras_client_remote_desktop_different_user.png")
        time.sleep(4 * self.dynamic_timing)

        pyautogui.write(self.ras_remote_user_name, interval=0.02 * self.dynamic_timing)
        time.sleep(4 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "22_ras_client_remote_desktop_user_name_type.png")

        pyautogui.press("tab")
        time.sleep(2 * self.dynamic_timing)

        pyautogui.write(self.ras_remote_password, interval=0.02 * self.dynamic_timing)      # Remote Desktop Password Type
        time.sleep(4 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "23_ras_client_remote_desktop_password_type.png")

        pyautogui.press("enter")        # Remote Desktop Connection
        time.sleep(5 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "24_ras_client_remote_desktop_connect.png")

        pyautogui.press("left")  # Remote Desktop Certificate Selection
        time.sleep(2 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "25_ras_client_remote_desktop_certificate_selection.png")

        pyautogui.press("enter")  # Remote Desktop Certificate Yes
        time.sleep(20 * self.dynamic_timing)

        # desktop_screen1_check_1, desktop_screen1_check_2 = self.img_obj.loop_retry_check(self.remote_screen1)
        desktop_screen2_check_1, desktop_screen2_check_2 = self.img_obj.loop_retry_check(self.remote_screen2)
        pyautogui.screenshot(self.logging_dir + "26_ras_client_remote_desktop_window.png")
        if desktop_screen2_check_1 is not None and desktop_screen2_check_2 is not None:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Successful")
            return True

        else:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Failed")
            sys.exit(0)

    def nat_act_connection(self):
        pyautogui.hotkey("alt", "ctrl", "shift", "e")
        time.sleep(10 * self.dynamic_timing)

        pyautogui.keyDown("alt")  # RASClient Tool Maximization
        pyautogui.keyDown("space")
        pyautogui.press("x")
        pyautogui.keyUp("alt")
        pyautogui.keyUp("space")
        time.sleep(4 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "27_ras_client_remote_desktop_window_browser.png")

        pyautogui.hotkey("ctrl", "t")  # Remote Desktop Browser New Tab
        time.sleep(6 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "28_ras_client_remote_desktop_window_browser_new_tab.png")

        pyautogui.write(self.ras_net_act_link, interval=0.02 * self.dynamic_timing)  # Remote Desktop Browser Report Link Type
        time.sleep(3 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "29_ras_client_remote_desktop_window_browser_link_type.png")

        pyautogui.press("enter")  # Remote Desktop Browser Report Link Page Load
        time.sleep(6 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "30_ras_client_remote_desktop_window_browser_link_page.png")

        # user_name_area_check_1, user_name_area_check_2 = self.img_obj.loop_retry_check(self.nat_link_opened)
        # pyautogui.click(user_name_area_check_1, clicks=2)
        pyautogui.click(813, 629)
        time.sleep(2 * self.dynamic_timing)

        pyautogui.write(self.ras_net_act_user_name, interval=0.02 * self.dynamic_timing)  # Remote Desktop Report User Name Type
        time.sleep(2 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "34_ras_client_remote_desktop_window_browser_user_name_type.png")

        pyautogui.press("tab")
        time.sleep(2 * self.dynamic_timing)

        pyautogui.write(self.ras_net_act_password, interval=0.02 * self.dynamic_timing)  # Remote Desktop Report Password Type
        time.sleep(2 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "35_ras_client_remote_desktop_window_browser_password_type.png")

        pyautogui.press("enter")  # Remote Desktop Connection Prompt
        time.sleep(10 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "36_ras_client_remote_desktop_window_browser_login_prompt.png")

        pyautogui.press("enter")  # Remote Desktop Report Screen
        time.sleep(10 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "37_ras_client_remote_desktop_window_browser_reporting_screen.png")

        report_folder_check_1, report_folder_check_2 = self.img_obj.loop_retry_check(self.report_folder)
        pyautogui.click(report_folder_check_1)     # Click on Report Folder
        pyautogui.screenshot(self.logging_dir + "38_ras_client_remote_desktop_window_browser_reporting_click.png")
        time.sleep(4 * self.dynamic_timing)

        performance_manager_check_1, performance_manager_check_2 = self.img_obj.loop_retry_check(self.performance_manager)
        pyautogui.click(performance_manager_check_1)  # Click On Performance Manager
        pyautogui.screenshot(self.logging_dir + "39_ras_client_remote_desktop_window_browser_performance_manager.png")
        time.sleep(6 * self.dynamic_timing)

        navigate_check_1, navigate_manager_check_2 = self.img_obj.loop_retry_check(self.navigate)
        pyautogui.click(navigate_check_1)  # Click On Navigate Button
        pyautogui.screenshot(self.logging_dir + "40_ras_client_remote_desktop_window_browser_navigate_button.png")
        time.sleep(4 * self.dynamic_timing)

        saved_reports_manager_check_1, saved_reports_manager_check_2 = self.img_obj.loop_retry_check(self.saved_reports)
        pyautogui.click(saved_reports_manager_check_1)  # Click On Saved Reports
        pyautogui.screenshot(self.logging_dir + "41_ras_client_remote_desktop_window_browser_saved_reports.png")
        time.sleep(4 * self.dynamic_timing)

        nat_act_connect_check_1, nat_act_connect_check_2 = self.img_obj.loop_retry_check(self.nat_act_connected)
        if nat_act_connect_check_1 is not None and nat_act_connect_check_2 is not None:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Successful")
            return True

        else:
            # self.mail_obj.shoot_mail(self.logging_dir + "5_ras_client_tool_maximization.png", "Failed")
            sys.exit(0)

    def report_download(self):

        pyautogui.click(103, 570)  # Click On Test Schedule Report
        time.sleep(5 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "43_ras_client_remote_desktop_window_browser_schedule_report.png")

        pyautogui.click(193, 597)  # Click On 2G Nokia Hourly
        time.sleep(10 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "44_ras_client_remote_desktop_window_browser_nokia_hourly_3G.png")

        export_button_check_1, export_button_check_2 = self.img_obj.loop_retry_check(self.export_button)
        pyautogui.click(export_button_check_1)  # Click On Export Button
        pyautogui.screenshot(self.logging_dir + "45_ras_client_remote_desktop_window_browser_export_button.png")
        time.sleep(4 * self.dynamic_timing)

        self.img_obj.loop_retry_check(self.ready_report)

        download_button_check_1, download_button_check_2 = self.img_obj.loop_retry_check(self.download_button)
        pyautogui.click(download_button_check_1)  # Click On Download Button
        pyautogui.screenshot(self.logging_dir + "45_ras_client_remote_desktop_window_browser_export_button.png")
        time.sleep(4 * self.dynamic_timing)

        pyautogui.hotkey("ctrl", "right")
        time.sleep(2 * self.dynamic_timing)
        pyautogui.write(".csv", interval=0.02 * self.dynamic_timing)
        time.sleep(2 * self.dynamic_timing)
        pyautogui.press("tab")
        time.sleep(1.5 * self.dynamic_timing)
        pyautogui.press("tab")
        time.sleep(1.5 * self.dynamic_timing)
        pyautogui.press("tab")
        time.sleep(1.5 * self.dynamic_timing)
        pyautogui.press("tab")
        time.sleep(1.5 * self.dynamic_timing)
        pyautogui.press("tab")
        time.sleep(1.5 * self.dynamic_timing)
        pyautogui.hotkey("enter")
        time.sleep(1.5 * self.dynamic_timing)
        pyautogui.write(self.report_directory, interval=0.02 * self.dynamic_timing)
        pyautogui.hotkey("enter")
        time.sleep(6 * self.dynamic_timing)
        pyautogui.screenshot(self.logging_dir + "46_ras_client_remote_desktop_window_browser_directory_path.png")

        directory_error_check_1, directory_error_check_2 = self.img_obj.image_check(self.directory_error)
        if directory_error_check_1 is None and directory_error_check_2 is None:
            pyautogui.hotkey("enter")
            pyautogui.hotkey("enter")

            self.img_obj.loop_retry_check(self.download_complete)
            pyautogui.screenshot(self.logging_dir + "47_ras_client_remote_desktop_window_browser_report_download_complete.png")

            pyautogui.hotkey("alt", "f4")  # Close Browser
            time.sleep(3 * self.dynamic_timing)
            pyautogui.press("left")
            time.sleep(2 * self.dynamic_timing)
            pyautogui.press("enter")
            time.sleep(6 * self.dynamic_timing)
            pyautogui.screenshot(self.logging_dir + "48_ras_client_remote_desktop_report_close_browser.png")

            desktop_directory_check_1, desktop_directory_check_1 = self.img_obj.loop_retry_check(self.desktop_directory)
            pyautogui.click(desktop_directory_check_1, clicks=2)
            time.sleep(4 * self.dynamic_timing)
            pyautogui.keyDown("alt")
            pyautogui.keyDown("space")
            pyautogui.press("x")
            pyautogui.keyUp("alt")
            pyautogui.keyUp("space")
            time.sleep(2 * self.dynamic_timing)
            pyautogui.screenshot(self.logging_dir + "49_ras_client_remote_desktop_report_directory_open.png")

            file_locate_check_1, file_locate_check_2 = self.img_obj.loop_retry_check(self.locate_file)
            x, y = pyautogui.center(file_locate_check_1)
            y = y + 38
            pyautogui.click(x, y)

            pyautogui.hotkey("ctrl", "c")
            time.sleep(2 * self.dynamic_timing)

            pyautogui.hotkey("win", "d")

            pyautogui.hotkey("ctrl", "v")
            time.sleep(1000000 * self.dynamic_timing)

        else:
            print("Incorrect Directory Path")
            sys.exit(0)
