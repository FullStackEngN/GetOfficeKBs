import datetime
import logging
import time


def download_file(browser, kb_number, product, target_folder):

    logger = logging.getLogger('download_kb')

    url = str.format('https://support.microsoft.com/help/{0}', kb_number)
    logger.info(str.format("Start download KB{0}: {1}", kb_number, url))
    
    browser.get(url)
    time.sleep(10)
   
    link_tag_32bit = ""
    link_tag_64bit = ""
    get_download_link_32bit = "false"
    get_download_link_64bit = "false"

    link_tag_list_32bit = []
    link_tag_list_32bit.append(str.format(
        "Download update {0} for 32-bit version of {1}", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download update KB{0} for 32-bit version of {1}", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download update KB{0} for 32-bit version of {1} for Office 2016", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download the 32-bit version of {0} update package now", product))
    link_tag_list_32bit.append(str.format(
        "Download security update {0} for the 32-bit version of {1}", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download security update KB{0} for the 32-bit version of {1}", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download security update KB {0} for the 32-bit version of {1}", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download the security update KB{0} for the 32-bit version of {1}", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download the security update KB{0} for 32-bit version of {1}", kb_number, product))
    link_tag_list_32bit.append(str.format(
        "Download the security update KB{0} for 32-bit version of {1}", kb_number, product))

    link_tag_list_64bit = []
    link_tag_list_64bit.append(str.format(
        "Download update {0} for 64-bit version of {1}", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download update KB{0} for 64-bit version of {1}", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download update KB{0} for 64-bit version of {1} for Office 2016", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download the 64-bit version of {0} update package now", product))
    link_tag_list_64bit.append(str.format(
        "Download security update {0} for the 64-bit version of {1}", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download security update KB{0} for the 64-bit version of {1}", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download security update KB {0} for the 64-bit version of {1}", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download the security update KB{0} for the 64-bit version of {1}", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download the security update KB{0} for 64-bit version of {1}", kb_number, product))
    link_tag_list_64bit.append(str.format(
        "Download the security update KB{0} for 64-bit version of {1}", kb_number, product))

    logging.debug("Try to get 32bit download link element")
    for item in link_tag_list_32bit:
        if(get_download_link_32bit == "false"):
            try:
                link_tag_32bit = browser.find_element_by_link_text(item)
                get_download_link_32bit = "true"

                logging.debug("Successfully get 32bit download link element")
            except:
                continue
        else:
            break
    
    logging.debug("Try to get 64bit download link element")
    for item in link_tag_list_64bit:
        if(get_download_link_64bit == "false"):
            try:
                link_tag_64bit = browser.find_element_by_link_text(item)
                get_download_link_64bit = "true"
                logging.debug("Successfully get 64bit download link element")
            except:
                continue
        else:
            break

    logging.debug("Try to get 32bit download link address")
    if(get_download_link_32bit == "true"):
        try:
            href_link_tag_32bit = link_tag_32bit.get_attribute("href")
            logger.info("32bit KB" + kb_number + ": " + href_link_tag_32bit)
        except:
            logger.error("Exception to get 32bit download link for this KB" + kb_number)
            get_download_link_32bit = "false"
            browser.save_screenshot(target_folder + "KB" + kb_number + "_32bit_warning.png")
    else:
        browser.save_screenshot(target_folder + "KB" + kb_number + "_32bit_warning.png")
        logger.warning("No matched text for 32bit KB" + kb_number)

    logging.debug("Try to get 64bit download link address")
    if(get_download_link_64bit == "true"):
        try:
            href_link_tag_64bit = link_tag_64bit.get_attribute("href")
            logger.info("64bit KB" + kb_number + ": " + href_link_tag_64bit)
        except:
            logger.error("Exception to get 64bit download link for this KB" + kb_number)
            get_download_link_64bit = "false"
            browser.save_screenshot(target_folder + "KB" + kb_number + "_64bit_warning.png")
    else:
        browser.save_screenshot(target_folder + "KB" + kb_number + "_64bit_warning.png")
        logger.warning("No matched text for 64bit KB" + kb_number)

    
    if(get_download_link_32bit == "true"):
        try:
            logger.info("******start download 32bit KB" + kb_number + " file")           
            logger.debug(href_link_tag_32bit)
            
            # download 32bit file
            browser.get(href_link_tag_32bit)
            
            #browser.implicitly_wait(20)
            time.sleep(30)

            download_link = browser.find_elements_by_class_name("download-button")
            
            logger.debug(len(download_link))
            logger.debug(download_link[0].get_attribute("href"))
            
            download_link_href = download_link[0].get_attribute("href")
            
            browser.get(download_link_href)

            time.sleep(30)
        except:
            logger.error("EXCEPTION TO DOWNLOAD 32bit KB" + kb_number + ": " + href_link_tag_32bit)
            browser.save_screenshot(target_folder + "KB" + kb_number + "_32bit_error.png")
    else:
        browser.save_screenshot(target_folder + "KB" + kb_number + "_32bit_error.png")
        logger.error("Don't find download link for 32bit KB" + kb_number + ", Stop the download process.")

    
    if(get_download_link_64bit == "true"):
        try:
            logger.info("******start download 64bit KB"+ kb_number +" file")
                            
            # download 64bit file
            logger.debug(href_link_tag_64bit)
            browser.get(href_link_tag_64bit)
            
            #browser.implicitly_wait(20)
            time.sleep(30)

            download_link = browser.find_elements_by_class_name("download-button")
            
            logger.debug(len(download_link))
            logger.debug(download_link[0].get_attribute("href"))
            
            download_link_href = download_link[0].get_attribute("href")
            browser.get(download_link_href)

            time.sleep(30)
        except:
            logger.error("EXCEPTION TO DOWNLOAD 64bit KB" + kb_number + ": " + href_link_tag_64bit)
            browser.save_screenshot(target_folder + "KB" + kb_number + "_64bit_error.png")
    else:
        browser.save_screenshot(target_folder + "KB" + kb_number + "_64bit_error.png")
        logger.error("Don't find download link for 64bit KB" + kb_number + ", Stop the download process.")

    logger.info("Finish download KB" + kb_number)