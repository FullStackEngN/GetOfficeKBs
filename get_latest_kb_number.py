# check whether there is duplicate kb number.
# find the component responding for the kb number.
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

current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + os.sep

FILENAME = current_script_folder + "log_" + os.path.basename(__file__) + ".log"

logger = create_logger_object(FILENAME)

logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

checked_kb_list = []
try:
    f = open(current_script_folder + "kb_list_checked.txt", "r")

    for line in f:
        if "," in line:
            tmp_list = line.split(",")

            for i in tmp_list:
                checked_kb_list.append(i)
        else:
            if line in checked_kb_list:
                logger.info("Duplicate KB number: " + line.strip().upper())
            else:
                checked_kb_list.append(line.strip().upper())

    logger.info("Read checked_kb_list file, length is: " + str(len(checked_kb_list)))
except:
    logger.info("Encounter exception when loading expected kb list.")

finally:
    f.close()

target_script_folder = current_script_folder + "Folder_Latest_KB_Numbers" + os.sep
logger.info("The target script folder is " + target_script_folder)

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


kb_component_list = []

browser = webdriver.Edge()

for kb in checked_kb_list:
    logger.info("KB number: " + kb.upper())

    kb_number = re.sub("KB", "", kb.upper())
    logger.info("KB number only: " + kb_number)

    kb_desc_url = "https://support.microsoft.com/kb/" + kb_number
    logger.info("KB description url: " + kb_desc_url)

    browser.get(kb_desc_url)
    # window_before = browser.window_handles[0]

    try:
        WebDriverWait(browser, 60).until(
            EC.visibility_of_element_located(
                By.XPATH, "//p[contains(text(),'-glb.exe')]"
            )
        )
    except:
        browser.save_screenshot(
            target_script_folder + "screenshots" + os.sep + kb_number + "_error.png"
        )
    finally:
        pass

    try:
        # file_hash_info_element = browser.find_element(
        #     "xpath", '//*[@id="ID0EDDBL"]').parent

        # file_hash_info_element = browser.find_element(
        #     By.XPATH, "//h3[contains(text(),'File hash information')]"
        # )

        # section_element = file_hash_info_element.find_element(By.XPATH, "./..")

        # package_name = section_element.find_element(
        #     By.XPATH, "//table/tbody/tr[1]/td[1]/p"
        # )

        package_name = browser.find_element(
            By.XPATH, "//p[contains(text(),'-glb.exe')]"
        )

        if package_name is not None:
            logger.info("Package name: " + package_name.text)

            kb_component_list.append(kb + "," + package_name.text)

        else:
            logger.error("Can't find 64bit or 32bit KB component name for " + kb_number)
            browser.save_screenshot(
                target_script_folder + "screenshots" + os.sep + kb_number + "_error.png"
            )
    except Exception as ex:
        browser.save_screenshot(
            target_script_folder + "screenshots" + os.sep + kb_number + "_error.png"
        )
        logger.info("Encounter exception: " + kb + ": " + str(ex))

    finally:
        pass

browser.quit()

kb_component_file = current_script_folder + "kb_list_component.txt"
with open(kb_component_file, "w") as f:
    for item in kb_component_list:
        f.write("%s\n" % item)

logger.info("The script ends.")
