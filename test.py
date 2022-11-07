#%%
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.common.by import By
from argparse import ArgumentParser
import glob
import os
import time
import re
# %%
parser = ArgumentParser()

parser.add_argument("--input")
parser.add_argument("--output")
parser.add_argument("--base")

args = parser.parse_args()

folder_script = args.input
result_path = args.output
base_url = args.base

if folder_script is None or result_path is None:
    print("Not Empty Folder input or output")

if not os.path.exists(folder_script):
    os.makedirs(folder_script)
if not os.path.exists(result_path):
    os.makedirs(result_path)
# %%
def find_action(content, bot):
    try:
        content_type = content.split(":")
        find_type = content_type[0].lower().strip()
        content_find = content_type[1].strip()
        if find_type == "xpath":
            element = bot.find_element(by=By.XPATH, value=content_find)
        elif find_type == "id":
            element = bot.find_element(by=By.ID, value=content_find)
        elif find_type == "class" or find_type == "classname":
            element = bot.find_element(by=By.CLASS_NAME, value=content_find)
        else:
            return (None, "Not Found type to find")
        return (element, "Success")
    except Exception as e:
        return (None, str(e['msg']))
#%%
def input_action(content, element):
    try:
        if element is None:
            return None
        element.send_keys(content)
        return (True, "Success")
    except Exception as e:
        return (False, str(e['msg']))

#%%
def get_action(url, bot):
    try:
        bot.get(url)
        return (True, "Success")
    except Exception as e:
        return (False, str(e['msg']))

# %%
def handle_automation(sheet, bot):
    for i in range(len(sheet)):
        type_action = sheet.iloc[i][0]
        if type_action == 'get':
            result_get = get_action(sheet.iloc[i][1], bot)
            if result_get[0] == False:
                break
            sheet['result'][i] = result_get[0]
            sheet['message'][i] = result_get[1]
        elif type_action == "find":
            elements = find_action(sheet.iloc[i][1], bot)
            element = elements[0]
            if element is None:
                sheet['result'][i] = False
            else:
                sheet['result'][i] = True
            sheet['message'][i] = elements[1]
        elif type_action == "input":
            result_input = input_action(sheet.iloc[i][1], element)
            sheet['result'][i] = result_input[0]
            sheet['message'][i] = result_input[1]
        elif type_action == "click":
            sheet['result'][i] = True
            sheet['message'][i] = "Success"
            element.click()
            time.sleep(2)
        elif type_action == "assert":
            result_find = find_action(sheet.iloc[i][1], bot)
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
                break
        elif type_action == "url":
            uri = re.sub(base_url, "", bot.current_url)
            if sheet.iloc[i][1] == uri:
                sheet['result'][i] = True
                sheet['message'][i] = "Success"
            else:
                sheet['result'][i] = False
                sheet['message'][i] = "Url Not Matched"
                break
    return sheet
# %%
def file_handle(input_path, result_path, file, bot):
    writer = pd.ExcelWriter(result_path, engine='openpyxl')
    print(file)
    for sheet_name in file.sheet_names:
        sheet = pd.read_excel(input_path, sheet_name=sheet_name)
        sheet['result'] = False
        sheet['message'] = "Not Execute"
        sheet = handle_automation(sheet, bot)
        sheet.to_excel(writer, sheet_name=sheet_name, index=None)
    writer.close()
# %%
files = glob.glob(f"./{folder_script}/*.xlsx")
# %%
bot = webdriver.Chrome()
# %%
for path in files:
    name = os.path.basename(path).split('.')[0]
    file_name_output = f'{result_path}/{name}_result.xlsx'
    print(name)
    file = pd.ExcelFile(path)
    file_handle(path, file_name_output, file, bot)
bot.close()




# %%

# %%
