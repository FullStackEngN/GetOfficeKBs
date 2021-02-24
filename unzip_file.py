import os
import shutil
from pathlib import Path
import subprocess

def run(cmd):
    ret = subprocess.call(cmd, shell=True)
    if ret != 0:
        print("Exec cmd " + cmd + " error, return value: " + str(ret))

source_folder = r"C:\Repos\GitHub\GetOfficeKBs\Office2016_KBs\x64"
target_folder = r"C:\Repos\GitHub\GetOfficeKBs\Office2016_KBs\x64_msp"
temp_folder = r"C:\Repos\GitHub\GetOfficeKBs\Office2016_KBs\temp"

if not os.path.exists(target_folder):
    os.makedirs(target_folder)

if not os.path.exists(temp_folder):
    os.makedirs(temp_folder)

#source_file = "C:\\Repos\\GitHub\\GetOfficeKBs\\Office2013_KBs\\x86\\KBKB2726958_fm20-x-none_940e36421ffcc83756b96d950b92b2985d0896a1.cab"
#target_folder = "C:\\Repos\\GitHub\\GetOfficeKBs\\Office2013_KBs\\x86_msp"
#cmd = "expand " + source_file + " -F:*.msp " + target_folder
# run(cmd)

for root, dirs, files in os.walk(source_folder):
    for name in files:
        if name.endswith("cab"):
            soure_file_path = root + os.sep + name
            print(soure_file_path)

            KB_number = name[:9]

            cmd = "expand " + soure_file_path + " -F:*.msp " + temp_folder
            run(cmd)

            for temproot, tempdirs, tempfiles in os.walk(temp_folder):
                for tempname in tempfiles:
                    temp_file_path = temproot + os.sep + tempname

                    target_file_name = tempname + "_" + \
                        KB_number + "_" + root[-3:] + ".msp"
                    target_file_path = target_folder + os.sep + target_file_name
                    print(target_file_path)

                    shutil.move(temp_file_path, target_file_path)
