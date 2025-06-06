import os
import pathlib
import wget
from selenium import webdriver

from extract_msp import extract_msp_from_cab
from get_msp_download_link import get_download_link
from my_logger_object import create_logger_object

current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + os.sep

FILENAME = os.path.join(current_script_folder, f"log_{os.path.basename(__file__)}.log")
logger = create_logger_object(FILENAME)

logger.info("The script starts running.")
logger.info(f"The script folder is {current_script_folder}")

target_download_folder = (
    os.path.join(current_script_folder, "Folder_Office2016_KBs") + os.sep
)
logger.info(f"The target download folder is {target_download_folder}")

if not os.path.exists(target_download_folder):
    os.makedirs(target_download_folder)
    logger.info("The target download folder doesn't exist, create it.")

kb_list_specified = []
try:
    with open(
        os.path.join(current_script_folder, "input_kb_list_specified.txt"), "r"
    ) as f:
        for line in f:
            kb_list_specified.append(line.strip().upper())
    logger.info(f"Read specified_kb_list file, length is: {len(kb_list_specified)}")
except Exception as ex:
    logger.info(f"No specified_kb_list file, so download all KBs. Exception: {ex}")

download_links = []

options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
browser = webdriver.Edge(options=options)

try:
    if len(kb_list_specified) > 0:
        logger.info(f"The specified KB list is: {kb_list_specified}")

        for current_kb_number in kb_list_specified:
            logger.info(f"Processing KB: {current_kb_number}")
            try:
                links = get_download_link(
                    browser, current_kb_number, target_download_folder
                )
                download_links += links
                logger.info(f">>>Done get links for {current_kb_number}")
            except Exception as ex:
                logger.error(f"<<<Failed to get links for {current_kb_number}: {ex}")

        with open(
            os.path.join(current_script_folder, "tmp_kb_links_list.log"), "w"
        ) as f:
            for element in download_links:
                f.write(f"{element[0]}:{element[1]}:{element[2]}\n")

        target_download_folder_x64 = os.path.join(target_download_folder, "x64")
        logger.info(
            f"The target download folder for 64bit KBs is {target_download_folder_x64}"
        )
        if not os.path.exists(target_download_folder_x64):
            os.makedirs(target_download_folder_x64)
            logger.info(
                "The target download folder for 64bit KBs doesn't exist, create it."
            )

        target_download_folder_x86 = os.path.join(target_download_folder, "x86")
        logger.info(
            f"The target download folder for 32bit KBs is {target_download_folder_x86}"
        )
        if not os.path.exists(target_download_folder_x86):
            os.makedirs(target_download_folder_x86)
            logger.info(
                "The target download folder for 32bit KBs doesn't exist, create it."
            )

        for item in download_links:
            try:
                tmp_file_name = wget.detect_filename(url=item[2])
                if item[0] == "x64":
                    out_path = os.path.join(
                        target_download_folder_x64, f"{item[1]}_{tmp_file_name}"
                    )
                elif item[0] == "x86":
                    out_path = os.path.join(
                        target_download_folder_x86, f"{item[1]}_{tmp_file_name}"
                    )
                else:
                    logger.warning(f"Unknown architecture: {item[0]}")
                    continue
                wget.download(url=item[2], out=out_path)
            except Exception as ex:
                logger.error(f"Failed to download {item}: {ex}")

        logger.info("Finish download progress.")
        logger.info("##############################")

        # Extract x86
        try:
            source_folder = target_download_folder_x86
            target_folder_msp = f"{target_download_folder_x86}_msp"
            temp_folder = f"{target_download_folder_x86}_temp"
            extract_msp_from_cab(source_folder, target_folder_msp, temp_folder)
        except Exception as ex:
            logger.error(f"Failed to extract x86 MSP: {ex}")

        # Extract x64
        try:
            source_folder = target_download_folder_x64
            target_folder_msp = f"{target_download_folder_x64}_msp"
            temp_folder = f"{target_download_folder_x64}_temp"
            extract_msp_from_cab(source_folder, target_folder_msp, temp_folder)
        except Exception as ex:
            logger.error(f"Failed to extract x64 MSP: {ex}")

    else:
        logger.info("No specified KB list, script exits.")

finally:
    browser.quit()
    logger.info("Browser closed.")

logger.info("Done...")
