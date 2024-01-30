import os
import requests
import zipfile

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FOLDER_NAME: str = 'automation_tool'
SOURCE_FOLDER = os.path.join(CURRENT_DIR, SOURCE_FOLDER_NAME)


def download_source():
    if os.path.exists(SOURCE_FOLDER):
        print('Already containing the source code')
        return

    download_url = f"https://github.com/hungdoan-tech/automation-tool/archive/main.zip"
    print("Start download source")
    response = requests.get(download_url)

    if response.status_code == 200:
        destination_directory = CURRENT_DIR
        file_name = os.path.join(destination_directory, "automation-tool.zip")
        with open(file_name, 'wb') as downloaded_zip_file:
            downloaded_zip_file.write(response.content)
        print("Download source successfully")

        with zipfile.ZipFile(file_name, 'r') as zip_ref:
            zip_ref.extractall(destination_directory)

        os.rename(os.path.join(destination_directory, 'automation-tool-main'),
                  os.path.join(destination_directory, SOURCE_FOLDER_NAME))
        os.remove(file_name)
        print(f"Extracted source code and placed it in {destination_directory}")
    else:
        print("Failed to download the source")


if __name__ == '__main__':
    download_source()
