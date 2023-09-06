import os
import pathlib
import shutil
from pathlib import Path

from my_logger_object import create_logger_object


def copy_component(component_kb_list, component_name, source_folder, target_folder):
    # source_folder = r"C:\CodeRepos\GetOfficeKBs\Folder_Office2016_KBs\x64_msp"
    # target_folder = r"C:\CodeRepos\GetOfficeKBs\Folder_Latest_KB_Numbers\x64_msp"
    # component_name = ""

    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    for root, dirs, files in os.walk(source_folder):
        for file_name in files:
            component_name_in_file = file_name.split("-")[0].strip()
            if component_name == component_name_in_file:    
                soure_file_path = root + os.sep + file_name
                target_file_path = target_folder + os.sep + file_name

                kb_number_in_file = file_name.split("_")[1].strip()

                if (component_name + "," + kb_number_in_file) not in component_kb_list:
                    component_kb_list.append(component_name + "," + kb_number_in_file)

                if os.path.isfile(soure_file_path):
                    try:
                        shutil.copy(soure_file_path, target_file_path)
                    except:
                        logger.debug("exception")


current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + os.sep

FILENAME = current_script_folder + "log_" + os.path.basename(__file__) + ".log"
logger = create_logger_object(FILENAME)

logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

component_list = []
try:
    f = open(current_script_folder + "output_msp_file_name_for_specified_kb.txt", "r")

    for line in f:
        component_str = line.split(",")[-1].strip()

        if component_str in component_list:
            logger.info("Duplicate component number: " + component_str)
        else:
            component_list.append(component_str)
except Exception as ex:
    logger.info("Encounter exception when loading expected kb list." + str(ex))
finally:
    f.close()

logger.info(len(component_list))

component_list.sort()
component_list_file = current_script_folder + "output_non_dup_component.txt"
with open(component_list_file, "w") as f:
    for item in component_list:
        f.write("%s\n" % item)

source_folder_x32 = r"C:\CodeRepos\GetOfficeKBs\Folder_Office2016_KBs\x86_msp"
target_folder_x32 = r"C:\CodeRepos\GetOfficeKBs\Folder_Latest_KB_Numbers\x86_msp"

source_folder_x64 = r"C:\CodeRepos\GetOfficeKBs\Folder_Office2016_KBs\x64_msp"
target_folder_x64 = r"C:\CodeRepos\GetOfficeKBs\Folder_Latest_KB_Numbers\x64_msp"


component_kb_list = []

for item in component_list:
    logger.debug(item)

    copy_component(component_kb_list, item, source_folder_x32, target_folder_x32)
    copy_component(component_kb_list, item, source_folder_x64, target_folder_x64)

component_kb_list.sort()

component_kb_list_file = current_script_folder + "output_latest_kb_for_component.txt"
with open(component_kb_list_file, "w") as f:
    for item in component_kb_list:
        f.write("%s\n" % item)


logger.info("Please check output file: " + component_kb_list_file)
logger.info("The script ends.")