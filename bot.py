import os
import re
import time
import psutil
import pyautogui
import subprocess
import win32com.client
from dmsService import Dms
from tulipService import Bot
from dataclasses import dataclass
from tulipService.model.variableModel import VariableModel
from rdp_service import BotRdpManager, RDPInputs


@dataclass
class BotInputSchema:
    language: VariableModel = None
    serverName: VariableModel = None
    client: VariableModel = None
    beekeeperUrl: VariableModel = None
    transaction_code: VariableModel = None
    dmscred: VariableModel = None
    sapCred: VariableModel = None
    rdp_cred_key: VariableModel = None
    login_cred_key: VariableModel = None
    server_address: VariableModel = None
    connect_rdp_api_url: VariableModel = None
    close_rdp_api_url: VariableModel = None
    access_token_baseurl: VariableModel = None
    server_name: VariableModel = None
    rebateAccountKeylumpsum: VariableModel = None
    rebateAgreementNumberLumpsum: VariableModel = None
    salesOrgVaiableLumpumsum: VariableModel = None
    materialNumber: VariableModel = None
    MaterialIterationValue: VariableModel = None
    materialAmount: VariableModel = None
    # waitingTime: VariableModel = None


@dataclass
class BotOutputSchema:
    def __init__(self):
        ...



class SAPGUI:
    def __init__(self, server, client, user, password, lang, logger):
        try:

            self.log = logger
            self.sap_gui_app = None
            self.application = None
            self.connection = None
            self.launch_sap_gui()

            self.session = self.connect_to_sap(server_name=server, client=client,
                                               username=user, password=password, language=lang)
        except Exception as e:
            self.log.error(f"exception from sap init method", e)
            raise e

    def is_process_running(self, name: str, retry_count=0, delay=0) -> bool:
        """
        Check if a process with the given name is running with retry logic.

        :param name: The name of the process to check.
        :param retry_count: Number of retry attempts.
        :param delay: Delay between retry attempts in seconds.
        :return: True if the process is running, False otherwise.
        """
        try:
            for attempt in range(retry_count + 1):
                try:
                    with subprocess.Popen(['tasklist', '/NH', '/FO', 'CSV'], stdout=subprocess.PIPE,
                                          encoding='utf-8') as tasklist:
                        for line in tasklist.stdout:
                            if name.lower() in line.lower():
                                self.log.info("Process found.")
                                return True
                except Exception as error:
                    self.log.error(f"An error occurred: {error}")

                if attempt < retry_count:
                    self.log.info("Retrying...")
                    time.sleep(delay)

            return False
        except Exception as e:
            self.log.error(f"exception from is process running method", e)
            raise e

    def terminate_process(self, name: str):
        """
        Terminates the process with the given name.

        :param name: The name of the process to terminate.
        """
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'].lower() == name.lower():
                    self.log.info(f"Terminating process {name} with PID {proc.info['pid']}.")
                    proc.terminate()
                    proc.wait(timeout=5)  # Wait for process to terminate
                    self.log.info(f"Process {name} terminated.")
                    return True
            self.log.info(f"No process with the name {name} found to terminate.")
            return False
        except Exception as error:
            self.log.error(
                f"An error occurred while terminating the process: {error}")
            return False

    # launching sapgui
    def launch_sap_gui(self):
        try:
            process_name = "saplogon.exe"
            self.log.info("start check process")
            if self.is_process_running(process_name, retry_count=3, delay=2):
                self.terminate_process(process_name)
                self.log.info("killing check process")
            self.log.info("after killing check process")
            os.startfile(process_name)
            # Wait for the SAP Logon process to be fully open and the SAPGUI object to be available
            retry_attempts = 8  # Adjust the number of retry attempts as needed
            sap_gui_available = False

            while retry_attempts > 0:
                if self.is_process_running(process_name, retry_count=1, delay=1):
                    try:
                        sap_gui = win32com.client.GetObject("SAPGUI")
                        if sap_gui is not None:
                            self.log.info("SAP GUI is available.")
                            sap_gui_available = True
                            break
                    except Exception as e:
                        self.log.info("Waiting for SAP GUI to become available...")
                else:
                    self.log.info("Waiting for SAP Logon to open...")

                time.sleep(2)
                retry_attempts -= 1

            if not sap_gui_available:
                self.log.error("SAP GUI did not become available within the expected time.")
        except Exception as e:
            self.log.error(f"exception from launch sap method", e)
            raise e

    # connect and login to sap
    def connect_to_sap(self, server_name, client, username, password, language):
        try:
            self.sap_gui_app = win32com.client.GetObject("SAPGUI")
            if not isinstance(self.sap_gui_app, win32com.client.CDispatch):
                return None

            application = self.sap_gui_app.GetScriptingEngine
            self.connection = application.OpenConnection(server_name, True)
            time.sleep(5)
            self.session = self.connection.Children(0)
            self.session.findById("wnd[0]").maximize()
            self.session.findById(
                "wnd[0]/usr/txtRSYST-BNAME").text = username
            self.session.findById(
                "wnd[0]/usr/pwdRSYST-BCODE").text = password
            self.session.findById(
                "wnd[0]/usr/txtRSYST-MANDT").text = client
            self.session.findById(
                "wnd[0]/usr/txtRSYST-LANGU").text = language
            self.session.findById("wnd[0]").sendVKey(0)
            return self.session
        except Exception as error:
            self.log.error(f"Error connecting to SAP: {error}")
            return error

    # Reading text from a text field
    def get_text_field(self, field_id):
        try:
            if self.connection:
                return self.session.findById(field_id).Text
        except Exception as error:
            self.log.error(f"Error getting text field: {error}")
            raise error

    def disconnect(self):
        try:
            if self.connection:
                self.session.CloseSession()
        except Exception as error:
            self.log.error("Error disconnecting from SAP: {error}")
            raise error

    def get_text_field(self, field_id):
        try:
            if self.connection:
                return self.session.findById(field_id).Text
        except Exception as error:
            self.log.error(f"Error getting text field: {error}")
            raise error


class BotLogic(Bot):
    def __init__(self) -> None:
        super().__init__()
        try:
            # Initialize an instance of BotOutputSchema
            self.outputs = BotOutputSchema()

            # Fetch the proposed bot inputs based on the schema
            self.input = self.bot_input.get_proposedBotInputs(BotInputs=BotInputSchema)
            self.vm_login = self.bot_input.get_identity(self.input.login_cred_key.value)

            self.rdp = self.bot_input.get_identity(self.input.rdp_cred_key.value)

            self.bot_inputs = RDPInputs(
                server_name=self.input.server_name.value,
                hex_id="some-guid-value",
                server_address=self.input.server_address.value,
                connect_rdp_api_url=self.input.connect_rdp_api_url.value,
                close_rdp_api_url=self.input.close_rdp_api_url.value,
                access_token_baseurl=self.input.access_token_baseurl.value,
                bot_class_name="Test"
            )

            self.rdp_connection = BotRdpManager(
                rdp_inputs=self.bot_inputs,
                oauth_credentials=self.rdp,
                vm_login_details=self.vm_login,
                logger=self.log
            )
            self.rdp_connection.connect_to_rdp()

            # Fetch the identity of the bot
            self._dms_identity = self.bot_input.get_identity(self.input.dmscred.value)
            self.sapidentity = self.bot_input.get_identity(self.input.sapCred.value)

            self.sap_gui = SAPGUI(server=self.input.serverName.value, client=self.input.client.value,
                                  user=self.sapidentity.credential.basicAuth.username,
                                  password=self.sapidentity.credential.basicAuth.password,
                                  lang=self.input.language.value,
                                  logger=self.log)

            self._dms = Dms(beekeeper_url=self.input.beekeeperUrl.value,
                            user_name=self._dms_identity.credential.basicAuth.username,
                            password=self._dms_identity.credential.basicAuth.password, logger=self.log)

        except Exception as error:
            # Log the error
            self.log.error(f"Error initializing BotLogic: {error}")
            self.bot_output.error(error)

    def dms_upload_screenshot(self, screenshot_path):
        """this function returns the signature of the file
        """
        try:
            self.log.info(f"dms screenshot path inside the functions {screenshot_path}")
            screenshot_fs = self._dms.upload_file_to_dms(complete_path=screenshot_path)

            self.log.info(f"dms screenshot path{screenshot_fs}")
            self.log.info("screenshot is uploaded in DMS")
            # self._dms.download_file_dms()
            return screenshot_fs
        except Exception as error:
            self.bot_output.error(error)
            raise f"Error occured during dms upload: {error}"

    def take_and_upload_screenshot(self, screenshot_name):
        """

        :param screenshot_name:
        :return: it passes the Scrennshot file into Dms
        """
        try:
            screenshot_filename = f"{screenshot_name}.png"
            full_screenshot_path = './'
            screenshot_path = os.path.join(full_screenshot_path, screenshot_filename)
            self.log.info(f"screenshot name: {screenshot_filename}")
            time.sleep(3)
            screenshot = pyautogui.screenshot()
            screenshot.save(screenshot_path)
            self.log.info(f"Screenshot saved successfully at {screenshot_path}.")
            screenshot_file = self.dms_upload_screenshot(screenshot_path=screenshot_path)

            self.log.info(f"Screenshot uploaded successfully: {screenshot_file}")
            return screenshot_file
        except Exception as error:
            self.log.error(f"Error taking and uploading screenshot: {error}")
            raise f"Error taking and uploading screenshot: {error}"
    def status_bar_error(self):
        try:
            # Access the status bar in the SAP session
            sbar = self.sap_gui.session.findById("wnd[0]/sbar")
            self.log.info("Checking the status bar for messages.")

            # Check if the status bar object exists and has text
            if sbar is not None and sbar.text != "":
                message = sbar.text
                self.log.info(f"Status bar message: {message}")

                # If the message indicates an error, return True
                if "no" in message.lower() or "documents" in message.lower() or "were" in message.lower() or "found" in message.lower():
                    self.bot_output.add_variable(key="dataFlag",
                                                 val="False")
                    return False
            self.bot_output.add_variable(key="dataFlag",
                                         val="True")
            return True
        except Exception as error:
            self.log.error(f"Error checking status bar: {error}")
            raise error

    def format_amount(self, amount):
        # Remove any existing dots and replace comma with dot
        amount = amount.replace(".", "").replace(",", ".")
        # Convert the string to a float
        float_amount = float(amount)
        # Format the float to the desired string format
        formatted_amount = "{:,.2f}".format(float_amount).replace(",", "X").replace(".", ",").replace("X", ".")
        return formatted_amount



    def main(self):
        """Bot Logic code"""
        try:
            self.log.info(f"transaction code: {self.input.transaction_code.value}")
            self.sap_gui.session.findById("wnd[0]/tbar[0]/okcd").text = self.input.transaction_code.value
            self.sap_gui.session.findById("wnd[0]").sendVKey(0)
            # Set the caret position in the input field
            self.sap_gui.session.findById(
                "wnd[0]/usr/ctxtRV13A-KNUMA_BO").text = self.input.rebateAgreementNumberLumpsum.value
            self.sap_gui.session.findById("wnd[0]/usr/ctxtRV13A-KNUMA_BO").caretPosition = 6

            # Send the Enter key
            self.sap_gui.session.findById("wnd[0]").sendVKey(0)

            # Press the button with index 23 on the toolbar
            self.sap_gui.session.findById("wnd[0]/tbar[1]/btn[23]").press()

            # Select the first row in the custom container shell
            self.sap_gui.session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").selectedRows = "0"

            # Double-click the current cell in the custom container shell
            self.sap_gui.session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").doubleClickCurrentCell()
            #
            # self.sap_gui.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,0]").setFocus()
            # self.sap_gui.session.findById(
            #     "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,0]").caretPosition = 10
            # materialNumber = self.sap_gui.session.findById(
            #     "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,0]").text
            # self.bot_output.add_variable(key="materialNumber", val=materialNumber)
            # # Set focus to the text field in the table
            # self.sap_gui.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").setFocus()
            #
            # # Set the caret position in the text field
            # self.sap_gui.session.findById(
            #     "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").caretPosition = 13
            # materialAmountUpdate = self.sap_gui.session.findById(
            #     "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").text
            #
            #
            # materialValue = materialAmountUpdate
            # materialValue = re.sub(r'-', '', materialValue)
            # self.bot_output.add_variable(key="materialAmountUpdate", val=materialValue)
            # print(materialValue)
            # screenshot = self.take_and_upload_screenshot(screenshot_name="vbo2MaterialAmountUpdateScreenshot")
            # self.bot_output.add_variable(key="vbo2MaterialAmountUpdateScreenshot",
            #                              val=screenshot)
            #
            # self.sap_gui.session.findById("wnd[0]/tbar[0]/okcd").text = self.input.transaction_code.value
            # self.sap_gui.session.findById("wnd[0]").sendVKey(0)
            # self.sap_gui.session.findById("wnd[0]/usr/ctxtRV13A-KNUMA_BO").text = self.input.rebateChangeAgreementNumberLumpsum.value
            # self.sap_gui.session.findById("wnd[0]/usr/ctxtRV13A-KNUMA_BO").caretPosition = 6
            # self.sap_gui.session.findById("wnd[0]").sendVKey(0)
            # self.sap_gui.session.findById("wnd[0]/tbar[1]/btn[9]").press()
            # self.sap_gui.session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").setCurrentCell(2, "GSTXT")
            # self.sap_gui.session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").selectedRows = "2"
            # self.sap_gui.session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shellcont/shell").doubleClickCurrentCell()
            # amount1 = self.input.materialAmount.value
            # amount2 = str(amount1)
            # formatted_amount1 = self.format_amount(amount2)  # Output: 6.500,00
            count = 0
            count1 = 1
            while count < int(self.input.MaterialIterationValue.value):
                count += 1
                count1 += 1
                self.sap_gui.session.findById(
                    "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,0]").caretPosition = 3
                materialNumber = self.sap_gui.get_text_field(
                    "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,0]")
                print(materialNumber)
                if int(materialNumber) == int(self.input.materialNumber.value):
                    # alreadyExistingMaterialAmount = self.sap_gui.session.findById(
                        # "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").text
                    # self.bot_output.add_variable(key="alreadyExistingMaterialAmount", val=alreadyExistingMaterialAmount)
                    materialamount = self.sap_gui.session.findById(
                        "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").text
                    materialValue = materialamount
                    # value = "13.46-"
                    if materialValue.endswith("-"):
                        materialValue = "-" + materialValue[:-1]
                    # print(value)
                    # materialValue = re.sub(r'-', '', materialValue)
                    self.bot_output.add_variable(key="materialAmountUpdate", val=materialValue)
                    # self.bot_output.add_variable(key="materialAmountUpdate", val=materialamount)
                    self.sap_gui.session.findById(
                        "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").setFocus()
                    self.sap_gui.session.findById(
                        "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[2,0]").caretPosition = 16
                    second_Screenshot = self.take_and_upload_screenshot(
                        screenshot_name="vbo2MaterialAmountUpdateScreenshot")
                    self.bot_output.add_variable(key="vbo2MaterialAmountUpdateScreenshot", val=second_Screenshot)
                    break
                self.sap_gui.session.findById(
                    "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").verticalScrollbar.position = count1
            # # self.sap_gui.session.findById(
            # #     "wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,0]").caretPosition = 10 # pointing the first row
            # # self.sap_gui.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").verticalScrollbar.position = 1
            # # self.sap_gui.session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").verticalScrollbar.position = 2
            #
            # self.sap_gui.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            # # waiting_time = int(self.input.waitingTime.value)
            # # time.sleep(waiting_time)
            self.bot_output.success()
        except Exception as e:
            # self.sap_gui.log.error(f"An error occurred during the process: {e}")
            self.log.error(f"Error in main execution: {e}")
            self.bot_output.error(e)
        finally:
            # if self.sap_gui:

            self.sap_gui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            self.sap_gui.session.findById("wnd[0]").sendVKey(0)
            self.rdp_connection.close_rdp()
