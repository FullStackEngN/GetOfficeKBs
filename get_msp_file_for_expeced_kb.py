import logging
import os
import pathlib
import wget
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from extract_msp import extract_msp_from_cab
from get_msp_download_link import get_download_link
from my_logger_object import create_logger_object


def get_kb_links_for_expected_kb(kb_number, browser, target_download_folder):
    links = []
    links = get_download_link(browser, kb_number, target_download_folder)

    return links


current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + os.sep

FILENAME = current_script_folder + "log_" + os.path.basename(__file__) + ".log"
logger = create_logger_object(FILENAME)

logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

target_download_folder = current_script_folder + "Folder_Office2016_KBs" + os.sep

if not os.path.exists(target_download_folder):
    os.makedirs(target_download_folder)
    logger.info("The target download folder doesn't exist, create it.")

kb_list_expected = []
try:
    f = open(current_script_folder + "kb_list_expected.txt", "r")

    for line in f:
        kb_list_expected.append(line.strip().upper())

    logging.info("Read kb_list_expected file, length is: " + str(len(kb_list_expected)))
except Exception as ex:
    logging.info("Encounter exception when loading expected kb list." + str(ex))
finally:
    f.close()

kb_list_download = []
kb_list_ignored = []
download_links = []

browser = webdriver.Firefox()

if len(kb_list_expected) > 0:
    for kb_number in kb_list_expected:
        kb_list_download.append(kb_number)

        current_kb_number = kb_number

        links = get_kb_links_for_expected_kb(
            current_kb_number, browser, target_download_folder
        )
        
        download_links += links

        logger.info("######Finish get links for target KB: " + current_kb_number)
else:
    logger.info("Target KBs are empty, Done.")

browser.quit()

logger.info("##############################")
logger.info("Start download progress.")

target_download_folder_x64 = target_download_folder + "x64"
logger.info("The target download folder for 64bit KBs is " + target_download_folder_x64)

if not os.path.exists(target_download_folder_x64):
    os.makedirs(target_download_folder_x64)
    logger.info("The target download folder for 64bit KBs doesn't exist, create it.")

target_download_folder_x86 = target_download_folder + "x86"
logger.info("The target download folder for 32bit KBs is " + target_download_folder_x86)

if not os.path.exists(target_download_folder_x86):
    os.makedirs(target_download_folder_x86)
    logger.info("The target download folder for 32bit KBs doesn't exist, create it.")

tmp_file_name = ""
for kb_number in download_links:
    if kb_number[0] == "x64":
        tmp_file_name = wget.detect_filename(url=kb_number[2])
        wget.download(
            url=kb_number[2],
            out=target_download_folder_x64
            + os.sep
            + kb_number[1]
            + "_"
            + tmp_file_name,
        )

    if kb_number[0] == "x86":
        tmp_file_name = wget.detect_filename(url=kb_number[2])
        wget.download(
            url=kb_number[2],
            out=target_download_folder_x86
            + os.sep
            + kb_number[1]
            + "_"
            + tmp_file_name,
        )

logger.info("Finish download progress.")
logger.info("##############################")

source_folder = target_download_folder_x86
target_folder = target_download_folder_x86 + "_" + "msp"
temp_folder = target_download_folder_x86 + "_" + "temp"

extract_msp_from_cab(source_folder, target_folder, temp_folder)

source_folder = target_download_folder_x64
target_folder = target_download_folder_x64 + "_" + "msp"
temp_folder = target_download_folder_x64 + "_" + "temp"

extract_msp_from_cab(source_folder, target_folder, temp_folder)

logger.info("Done...")
