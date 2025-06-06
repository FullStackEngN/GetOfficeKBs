import os
import shutil
import subprocess
from pathlib import Path


def run(cmd):
    try:
        result = subprocess.run(cmd, shell=True, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Exec cmd '{cmd}' error, return value: {e.returncode}")


def extract_msp_from_cab(source_folder, target_folder, temp_folder):
    source_folder = Path(source_folder)
    target_folder = Path(target_folder)
    temp_folder = Path(temp_folder)

    target_folder.mkdir(parents=True, exist_ok=True)
    temp_folder.mkdir(parents=True, exist_ok=True)

    for root, dirs, files in os.walk(source_folder):
        for name in files:
            if name.lower().endswith(".cab"):
                source_file_path = Path(root) / name
                print(source_file_path)

                KB_number = name[:9]

                cmd = f'expand "{source_file_path}" -F:*.msp "{temp_folder}"'
                run(cmd)

                for temproot, tempdirs, tempfiles in os.walk(temp_folder):
                    for tempname in tempfiles:
                        temp_file_path = Path(temproot) / tempname
                        target_file_name = f"{tempname}_{KB_number}_{Path(root).name[-3:]}.msp"
                        target_file_path = target_folder / target_file_name
                        print(target_file_path)
                        try:
                            shutil.move(str(temp_file_path), str(target_file_path))
                        except Exception as e:
                            print(f"Failed to move {temp_file_path} to {target_file_path}: {e}")

    if temp_folder.exists() and temp_folder.is_dir():
        try:
            shutil.rmtree(temp_folder)
        except Exception as e:
            print(f"Failed to remove temp folder {temp_folder}: {e}")
