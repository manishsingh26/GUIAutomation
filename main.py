
import os
import sys
root_path_ = r"C:\Users\m4singh\PycharmProjects\ReportDownloadRASClient\\"
sys.path.append(root_path_)
import json
from Script.utils.system_check import CheckOperation
from Script.automation.tool_main import AutomationActivity


class Operation(object):

    def __init__(self, root_path):
        self.app_config_path = root_path + "Config/app_config.json"
        with open(self.app_config_path) as config:
            self.app_config = json.load(config)
        self.app_config["root_path"] = root_path
        self.app_config["logging_path"] = root_path + self.app_config["logging_path"] + os.path.sep

    def check_operation(self):
        check_obj = CheckOperation(self.app_config)
        check_obj.resolution_check()
        check_obj.caps_lock_check()
        check_obj.ras_software_check()
        check_obj.ram_check()

    def ras_operation(self):
        auto_obj = AutomationActivity(self.app_config)
        login_status = auto_obj.ras_software_login()
        if login_status is True:
            connection_status = auto_obj.customer_config()
            if connection_status is True:
                remote_connection_status = auto_obj.remote_connection()
                if remote_connection_status is True:
                    nat_act_connection_status = auto_obj.nat_act_connection()
                    if nat_act_connection_status is True:
                        auto_obj.report_download()


if __name__ == "__main__":

    obj = Operation(root_path_)
    # obj.check_operation()
    obj.ras_operation()
