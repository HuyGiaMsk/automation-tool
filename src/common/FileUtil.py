import copy
import os
import re
import zipfile
from datetime import datetime, timedelta
from logging import Logger
from typing import Callable

import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger
from src.setup.packaging.path.PathResolvingService import PathResolvingService


def persist_settings_to_file(task_name: str, setting_values: dict[str, str]):
    input_dir: str = PathResolvingService.get_instance().get_input_dir()
    file_path: str = os.path.join(input_dir, "{}.properties".format(task_name))
    with ResourceLock(file_path=file_path):
        with open(file_path, 'w') as file:
            file.truncate(0)

        with open(file_path, 'a') as file:
            for key, value in setting_values.items():
                file.write(f"{key} = {value}\n")

    logger = get_current_logger()
    logger.debug("Data persisted successfully")


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

    result = re.search(r'([a-zA-Z]+)(\d+)', start_cell)
    if result:
        column = result.group(1)
        start_row = int(result.group(2))
    else:
        raise Exception("Not match excel cell position format")

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


def find_module(current_path: str, task_name: str) -> str | None:
    for item in os.listdir(current_path):

        item_path = os.path.join(current_path, item)

        if os.path.isdir(item_path):
            matched_file = find_module(item_path, task_name)
            if matched_file is None:
                continue

            return matched_file

        if item == f"{task_name}.py":
            path_parts: list[str] = item_path.rsplit(os.sep)

            is_collecting: bool = False
            real_module_paths: list[str] = []
            for part in path_parts:

                if is_collecting:
                    real_module_paths.append(part)
                    continue

                if part == 'src':
                    real_module_paths.append(part)
                    is_collecting = True
                    continue

            final_module_path: str = ".".join(real_module_paths)
            final_module_path = final_module_path[:-3]
            return final_module_path

    return None


def get_all_concrete_task_names() -> list[str]:
    discarded_task_names: set[str] = {"__init__", "AutomatedTask", "DesktopTask", "WebTask"}
    concrete_task_names: set[str] = get_files_names_in_dir(dir_path=PathResolvingService.get_instance().get_task_dir(),
                                                           file_extension='.py',
                                                           excluded_file_names=discarded_task_names)
    return sorted(list(concrete_task_names))


def get_files_names_in_dir(dir_path: str, file_extension: str, excluded_file_names: set[str]) -> set[str]:
    logger: Logger = get_current_logger()

    if file_extension is None:
        file_extension = ".py"

    file_names: set[str] = set()
    for root, _, files in os.walk(dir_path):
        for file in files:
            logger.warning(file)
            if file.lower().endswith(file_extension):
                clean_name = file.replace(".py", "")
                file_names.add(clean_name)

    if excluded_file_names is None or excluded_file_names.__len__() == 0:
        return file_names

    return file_names - excluded_file_names
