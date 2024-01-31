import re
import copy
from typing import Callable

import openpyxl
import zipfile
import os

from datetime import datetime, timedelta
from logging import Logger
from openpyxl.cell.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from src.common.ThreadLocalLogger import get_current_logger
from src.common.ResourceLock import ResourceLock


def load_key_value_from_file_properties(setting_file: str) -> dict[str, str]:
    logger: Logger = get_current_logger()
    if not os.path.exists(setting_file):
        raise Exception("The settings file {} is not existed. Please providing it !".format(setting_file))

    settings: dict[str, str] = {}
    logger.info('Start loading settings from file')

    with ResourceLock(file_path=setting_file):

        with open(setting_file, 'r') as setting_file_stream:

            for line in setting_file_stream:

                line = line.replace("\n", "").strip()
                if len(line) == 0 or line.startswith("#"):
                    continue

                key_value: list[str] = line.split("=")
                key: str = key_value[0].strip()
                value: str = key_value[1].strip()
                settings[key] = value

        return settings


def get_excel_data_in_column_start_at_row(file_path, sheet_name, start_cell) -> list[str]:
    logger: Logger = get_current_logger()
    column: str = 'A'
    start_row: int = 0

    result = re.search('([a-zA-Z]+)(\d+)', start_cell)
    if result:
        column = result.group(1)
        start_row = int(result.group(2))

    file_path = r'{}'.format(file_path)
    logger.info(
        r"Read data from file {} at sheet {}, collect all data at column {} start from row {}".format(
            file_path, sheet_name, column, start_row))

    with ResourceLock(file_path=file_path):

        workbook: Workbook = openpyxl.load_workbook(filename=r'{}'.format(file_path), data_only=True, keep_vba=True)
        worksheet: Worksheet = workbook[sheet_name]

        values: list[str] = []
        runner: int = 0
        max_index = start_row - 1
        for cell in worksheet[column]:
            cell: Cell = cell

            if runner < max_index:
                runner += 1
                continue

            if cell.value is None:
                continue

            values.append(str(cell.value))
            runner += 1

        if len(values) == 0:
            logger.error(
                r'Not containing any data from file {} at sheet {} at column {} start from row {}'.format(
                    file_path, sheet_name, column, start_row))
            raise Exception("Not containing required data in the specified place in file Excel")

        logger.info('Collect data from excel file successfully')
        return values


def extract_zip(zip_file_path: str,
                extracted_dir: str,
                callback_on_root_folder: Callable[[str], None],
                callback_on_extracted_folder: Callable[[str], None]) -> None:
    logger: Logger = get_current_logger()
    if not os.path.isfile(zip_file_path) or not zip_file_path.lower().endswith('.zip'):
        raise Exception('{} is not a zip file'.format(zip_file_path))

    zip_file_path = zip_file_path.replace('\\', '/')
    file_name_contain_extension: str = zip_file_path.split('/')[-1]
    clean_file_name: str = file_name_contain_extension.split('.')[0]

    if not extracted_dir.endswith('/') and not extracted_dir.endswith('\\'):
        extracted_dir += '\\'

    if clean_file_name not in extracted_dir:
        extracted_dir = r'{}{}'.format(extracted_dir, clean_file_name)
        if not os.path.exists(extracted_dir):
            os.mkdir(extracted_dir)

    logger.debug(r'Start extracting file {} into {}'.format(zip_file_path, extracted_dir))

    with ResourceLock(file_path=zip_file_path):
        with ResourceLock(file_path=extracted_dir):
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(extracted_dir)

    if callback_on_extracted_folder is not None:
        callback_on_extracted_folder(extracted_dir)

    os.remove(zip_file_path)
    logger.debug(r'Extracted successfully file {} to {}'.format(zip_file_path, extracted_dir))

    if callback_on_root_folder is not None:
        current_dir: str = os.path.dirname(os.path.abspath(zip_file_path))
        callback_on_root_folder(current_dir)


def check_parent_folder_contain_all_required_sub_folders(parent_folder: str,
                                                         required_sub_folders: set[str]) -> (bool, set[str], set[str]):
    contained_set: set[str] = set()
    with ResourceLock(file_path=parent_folder):

        for dir_name in os.listdir(parent_folder):
            full_dir_name = os.path.join(parent_folder, dir_name)
            if os.path.isdir(full_dir_name) and dir_name in required_sub_folders:
                contained_set.add(dir_name)
                required_sub_folders.discard(dir_name)

        is_all_contained: bool = len(required_sub_folders) == 0
        not_contained_set: set[str] = copy.deepcopy(required_sub_folders)
        return is_all_contained, contained_set, not_contained_set


def remove_all_in_folder(folder_path: str,
                         only_files: bool = False,
                         file_extension: str = None,
                         elapsed_time: timedelta = None) -> None:
    logger: Logger = get_current_logger()
    with ResourceLock(file_path=folder_path):

        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                if file_extension is None:
                    os.remove(file_path)
                    continue

                if not file_path.endswith(file_extension):
                    continue

                since_datetime = datetime.now() - elapsed_time
                if datetime.fromtimestamp(os.path.getctime(file_path)) > since_datetime:
                    os.remove(file_path)
                    continue
            else:
                if not only_files:
                    remove_all_in_folder(file_path)
        logger.debug(
            r'Deleted all {} in folder {}'.format(file_extension, folder_path))
