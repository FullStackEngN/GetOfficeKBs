# check whether there is duplicate kb number.
# find the component corresponding for the kb number.
# get latest update for the component.
# compare with previous kb number, if same, just download. if not, download latest one.

import os
import pathlib
import re

from my_logger_object import create_logger_object

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def add_kb_to_list(logger, checked_kb_list, kb_number_str):
    if kb_number_str in checked_kb_list:
        logger.info("Duplicate KB number: " + kb_number_str)
    else:
        checked_kb_list.append(kb_number_str)


current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + os.sep

FILENAME = current_script_folder + "log_" + os.path.basename(__file__) + ".log"
logger = create_logger_object(FILENAME)

logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

target_script_folder = current_script_folder + "Folder_Latest_KB_Numbers" + os.sep
logger.info("The target folder is " + target_script_folder)

if not os.path.exists(target_script_folder):
    os.makedirs(target_script_folder)
    logger.info("The target script folder doesn't exist, create it.")

if not os.path.exists(target_script_folder + "screenshots"):
    os.makedirs(target_script_folder + "screenshots")
    logger.info(
        "Screenshots folder doesn't exist, create it: "
        + target_script_folder
        + "screenshots"
    )

checked_kb_list = []

try:
    f = open(current_script_folder + "kb_list_checked.txt", "r")

    for line in f:
        kb_number_str = None

        if "," in line:
            tmp_list = line.split(",")

            for i in tmp_list:
                kb_number_str = i.strip().upper()

                if len(kb_number_str) == 0:
                    continue

                add_kb_to_list(logger, checked_kb_list, kb_number_str)
        else:
            kb_number_str = line.strip().upper()

            if len(kb_number_str) == 0:
                continue

            add_kb_to_list(logger, checked_kb_list, kb_number_str)

    logger.info("Read checked_kb_list file, length is: " + str(len(checked_kb_list)))
except:
    logger.info("Encounter exception when loading expected kb list.")
finally:
    f.close()

kb_component_list = []
kb_component_error_list = []

browser = webdriver.Edge()

for kb in checked_kb_list:
    kb_str = kb.strip().upper()

    logger.info("KB number: " + kb_str)

    kb_number = re.sub("KB", "", kb_str)
    logger.info("KB number only: " + kb_number)

    kb_desc_url = "https://support.microsoft.com/kb/" + kb_number
    logger.info("KB description url: " + kb_desc_url)

    browser.get(kb_desc_url)

    package_name = None

    try:
        package_name = WebDriverWait(browser, 30).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//p[contains(text(),'-glb.exe')]")
            )
        )
    except:
        try:
            package_name = WebDriverWait(browser, 30).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//span[contains(text(),'-glb.exe')]")
                )
            )
        except Exception as ex:
            browser.save_screenshot(
                target_script_folder + "screenshots" + os.sep + kb_str + "_error.png"
            )

            logger.info(
                "Encounter exception when waiting/loading element for "
                + kb_str
                + " : "
                + str(ex)
            )

            kb_component_error_list.append(kb_str + ", ")

    if package_name is not None:
        logger.info("Package name: " + package_name.text)

        kb_component_list.append(kb_str + "," + package_name.text)
    else:
        logger.error("Can't find component name for " + kb_str)
        browser.save_screenshot(
            target_script_folder + "screenshots" + os.sep + kb_str + "_error.png"
        )

browser.quit()

logger.info("kb_list_component count: " + str(len(kb_component_list)))
logger.info("kb_list_component_error count: " + str(len(kb_component_error_list)))

kb_component_file = current_script_folder + "kb_list_component.txt"
with open(kb_component_file, "w") as f:
    for item in kb_component_list:
        f.write("%s\n" % item)

kb_component_error_file = current_script_folder + "kb_list_component_error.txt"
with open(kb_component_error_file, "w") as f:
    for item in kb_component_error_list:
        f.write("%s\n" % item)

logger.info("The script ends.")
