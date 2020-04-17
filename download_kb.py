import datetime
import time


def DownloadFile(browser, KB, product, security):

    browser.get(str.format('https://support.microsoft.com/help/{0}', KB))
    browser.implicitly_wait(5)

    link_tag_32bit = ""
    link_tag_64bit = ""

    if(security == "true"):
        link_tag_32bit = browser.find_element_by_link_text(
            str.format("Download security update {0} for the 32-bit version of {1}", KB, product))
        link_tag_64bit = browser.find_element_by_link_text(
            str.format("Download security update {0} for the 64-bit version of {1}", KB, product))
    else:
        link_tag_32bit = browser.find_element_by_link_text(
            str.format("Download update KB{0} for 32-bit version of {1}", KB, product))
        link_tag_64bit = browser.find_element_by_link_text(
            str.format("Download update KB{0} for 64-bit version of {1}", KB, product))

    href_link_tag_32bit = link_tag_32bit.get_attribute("href")
    href_link_tag_64bit = link_tag_64bit.get_attribute("href")

    print("start download 32bit KB file")
    # download 32bit file
    print(href_link_tag_32bit)
    browser.get(href_link_tag_32bit)
    browser.implicitly_wait(5)

    download_link = browser.find_elements_by_class_name("download-button")
    # print(len(download_link))
    # print(download_link[0].get_attribute("href"))
    download_link_href = download_link[0].get_attribute("href")
    browser.get(download_link_href)

    #currentDT = datetime.datetime.now()
    #print(str(currentDT))

    # browser.implicitly_wait(120)
    time.sleep(15)

    #currentDT = datetime.datetime.now()
    #print(str(currentDT))

    print("start download 64bit KB file")
    # download 64bit file
    print(href_link_tag_64bit)
    browser.get(href_link_tag_64bit)
    browser.implicitly_wait(5)

    download_link = browser.find_elements_by_class_name("download-button")
    # print(len(download_link))
    # print(download_link[0].get_attribute("href"))
    download_link_href = download_link[0].get_attribute("href")
    browser.get(download_link_href)
    time.sleep(15)
