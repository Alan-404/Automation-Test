from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.common.by import By
import glob
import os
import time


class BotTest:
    def __init__(self, input_path=None, result_path=None):
        self.bot = webdriver.Chrome()
        self.input_path = input_path
        self.result_path = result_path

    def find_action(self, content):
        try:
            content_type = content.split(":")
            find_type = content_type[0].lower().strip()
            content_find = content_type[1].strip()
            if find_type == "xpath":
                element = self.bot.find_element(by=By.XPATH, value=content_find)
            elif find_type == "id":
                element = self.bot.find_element(by=By.ID, value=content_find)
            elif find_type == "class" or find_type == "classname":
                element = self.bot.find_element(by=By.CLASS_NAME, value=content_find)
            else:
                return (None, "Not Found type to find")
            return (element, "Success")
        except Exception as e:
            return (None, str(e['msg']))

    def input_action(self, content, element):
        try:
            if element is None:
                return (False, "Not Found Element")
            element.send_keys(content)
            return (True, "Success")
        except Exception as e:
            return (False, str(e['msg']))
    
    def get_action(self, url):
        try:
            self.bot.get(url)
            return (True, "Success")
        except Exception as e:
            return (False, str(e['msg']))
    
    def handle_automation(self, sheet):
        for i in range(len(sheet)):
            type_action = sheet.iloc[i][0]
            if type_action == 'get':
                result_get = self.get_action(sheet.iloc[i][1], self.bot)
                if result_get[0] == False:
                    break
                sheet['result'][i] = result_get[0]
                sheet['message'][i] = result_get[1]
            elif type_action == "find":
                elements = self.find_action(sheet.iloc[i][1], self.bot)
                element = elements[0]
                if element is None:
                    sheet['result'][i] = False
                else:
                    sheet['result'][i] = True
                sheet['message'][i] = elements[1]
            elif type_action == "input":
                result_input = self.input_action(sheet.iloc[i][1], element)
                sheet['result'][i] = result_input[0]
                sheet['message'][i] = result_input[1]
            elif type_action == "click":
                sheet['result'][i] = True
                sheet['message'][i] = "Success"
                element.click()
                time.sleep(2)
            elif type_action == "assert":
                result_find = self.find_action(sheet.iloc[i][1], self.bot)
                element = result_find[0]
                if result_find[0] is None:
                    sheet['result'][i] = False
                    sheet['message'][i] = "Not Found Element"
                    break
                if element.text.strip() == sheet['text'][i].strip():
                    sheet['result'][i] = True
                    sheet['message'][i] = "Success"
                else:
                    sheet['result'][i] = False
                    sheet['message'][i] = "Not Matched"
        return sheet

    def file_handle(self, input_path, file, result_path):
        writer = pd.ExcelWriter(result_path, engine='openpyxl')
        for sheet_name in file.sheet_names:
            sheet = pd.read_excel(input_path, sheet_name=sheet_name)
            sheet['result'] = False
            sheet['message'] = "Not Execute"
            sheet = self.handle_automation(sheet)
            sheet.to_excel(writer, sheet_name=sheet_name, index=None)
        writer.close()

    def test(self):
        files = glob.glob(f"{self.input_path}/*.xlsx")
        for path in files:
            name = os.path.basename(path).split('.')[0]
            file_name_output = f'{self.result_path}/{name}_result.xlsx'
            file = pd.ExcelFile(path)
            self.file_handle(path, file_name_output, file)
        self.bot.close()
        print("========Done==================")
