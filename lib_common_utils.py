import os

import wget


def read_kb_list(folder, filename, logger):
    """
    Read a list of KB numbers from a file and return them as a list.
    The KB numbers are converted to uppercase and stripped of whitespace.
    If the file does not exist, return an empty list.
    """

    kb_list = []

    path = os.path.join(folder, filename)
    if os.path.exists(path):
        with open(path, "r") as f:
            for line in f:
                kb_list.append(line.strip().upper())
        logger.info(f"Read {filename}, length is: {len(kb_list)}")
    else:
        logger.info(f"No {filename} file, so skip.")
    return kb_list


def ensure_folder_exists(folder_path, logger=None):
    """
    Ensure that the specified folder exists. If it does not exist, create it.
    """

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        if logger:
            logger.info(f"Created folder: {folder_path}")


def check_kb_in_excluded_list(current_kb_number, kb_list_excluded, logger):
    if current_kb_number in kb_list_excluded:
        logger.info(f">>> @@@ ---{current_kb_number} is excluded by user")
        return True
    return False


def download_file(arch, filename, url, target_folder, logger):
    """
    Download a file and save it to the target folder with a specific naming convention.
    """

    try:
        detected_name = wget.detect_filename(url=url)
        out_path = os.path.join(target_folder, f"{filename}_{detected_name}")
        wget.download(url=url, out=out_path)
        logger.info(f"Downloaded {out_path}")
    except Exception as ex:
        logger.error(f"Failed to download {arch}, {filename}, {url}: {ex}")


def process_download_links(download_links, target_download_folder_x64, target_download_folder_x86, logger):
    """
    Process download links for both x64 and x86 architectures.
    """
    for link_info in download_links:
        arch, filename, url = link_info
        if arch == "x64":
            download_file("x64", filename, url, target_download_folder_x64, logger)
        elif arch == "x86":
            download_file("x86", filename, url, target_download_folder_x86, logger)
        else:
            logger.warning(f"Unknown architecture: {arch}")


def save_screenshot_with_log(browser, path, logger, msg=None, exception=False):
    """Helper to save screenshot and log message."""
    browser.save_screenshot(path)
    if msg:
        if exception:
            logger.exception(msg)
        else:
            logger.error(msg)
