# This script automates the process of downloading Office KB files (CAB format)
# It reads KB numbers from an input file: input_kb_list_expected.txt.
# Uses Selenium to retrieve download links, downloads the files into architecture-specific folders,
# Then extracts MSP files from downloaded CABs.

import os
import pathlib

import wget
from selenium import webdriver

from lib_common_utils import ensure_folder_exists, read_kb_list
from lib_extract_msp import extract_msp_from_cab
from lib_get_msp_download_link import get_download_link
from lib_logger_object import create_logger_object


def download_kb_files(download_links, target_folder_x64, target_folder_x86, logger):
    """
    Download KB files to the corresponding architecture folders.
    """
    for kb_info in download_links:
        arch, kb_number, url = kb_info
        tmp_file_name = wget.detect_filename(url=url)
        if arch == "x64":
            out_path = os.path.join(target_folder_x64, f"{kb_number}_{tmp_file_name}")
        elif arch == "x86":
            out_path = os.path.join(target_folder_x86, f"{kb_number}_{tmp_file_name}")
        else:
            logger.info(f"Unknown architecture for KB: {kb_info}")
            continue
        wget.download(url=url, out=out_path)
        logger.info(f"Downloaded {arch} KB: {kb_number} to {out_path}")


def main():
    current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + os.sep
    log_file_path = os.path.join(current_script_folder, f"log_{os.path.basename(__file__)}.log")
    logger = create_logger_object(log_file_path)

    logger.info("Script started.")
    logger.info(f"Script folder: {current_script_folder}")

    target_download_folder = os.path.join(current_script_folder, "Folder_Office2016_KBs")
    ensure_folder_exists(target_download_folder, logger)

    kb_list_file = os.path.join(current_script_folder, "input_kb_list_expected.txt")
    kb_list_expected = read_kb_list(kb_list_file)

    download_links = []

    if kb_list_expected:
        options = webdriver.EdgeOptions()
        options.add_argument("--start-maximized")
        browser = webdriver.Edge(options=options)
        try:
            for kb_number in kb_list_expected:
                links = get_download_link(browser, kb_number, target_download_folder)
                download_links.extend(links)
                logger.info(f"Finished getting links for KB: {kb_number}")
        finally:
            browser.quit()
    else:
        logger.info("No target KBs found. Exiting.")
        return

    logger.info("##############################")
    logger.info("Start download process.")

    target_download_folder_x64 = os.path.join(target_download_folder, "x64")
    target_download_folder_x86 = os.path.join(target_download_folder, "x86")
    ensure_folder_exists(target_download_folder_x64, logger)
    ensure_folder_exists(target_download_folder_x86, logger)

    download_kb_files(download_links, target_download_folder_x64, target_download_folder_x86, logger)

    logger.info("Download process finished.")
    logger.info("##############################")

    # Extract MSP files from CABs for both architectures
    for arch_folder in [target_download_folder_x86, target_download_folder_x64]:
        source_folder = arch_folder
        target_folder = f"{arch_folder}_msp"
        temp_folder = f"{arch_folder}_temp"
        extract_msp_from_cab(source_folder, target_folder, temp_folder)

    logger.info("All done.")


if __name__ == "__main__":
    main()
