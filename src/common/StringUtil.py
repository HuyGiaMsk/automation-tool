import os.path
import re
from logging import Logger
from re import Match
from typing import Tuple

from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger
from src.setup.packaging.path.PathResolvingService import PathResolvingService


def validate_keys_of_dictionary(settings: dict[str, str],
                                mandatory_settings: set[str]) -> None:
    logger: Logger = get_current_logger()
    error_messages: list[str] = []
    for key in mandatory_settings:

        value: str = settings.get(str(key))
        if value is None:
            error_messages.append('{} is missing'.format(key))

        if len(value) == 0:
            error_messages.append('{} is missing'.format(key))

    if len(error_messages) != 0:
        logger.error('Invalid value at key : ')
        for error in error_messages:
            logger.error(error)
        raise Exception('Invalid settings!')


def decode_url(url: str) -> str:
    decoded: str = ""
    index: int = 0

    while index < len(url):
        if url[index] == '%':
            hex_char = url[index + 1: index + 3]  # Extract the hexadecimal representation of the character
            decoded += chr(int(hex_char, 16))  # Convert it to an integer and then to a character
            index += 3  # Skip the '%' and two hexadecimal digits
        else:
            decoded += url[index]
            index += 1

    return decoded


def escape_bat_file_special_chars(input_file: str = '.\\Downloadsrc.py',
                                  output_file: str = os.path.join(PathResolvingService.get_instance().get_output_dir(),
                                                                  'EscapedCharsEmbeddedPythonToBat.output')) -> None:
    if not os.path.exists(input_file):
        raise Exception('invalid input')

    if os.path.exists(output_file):
        os.remove(output_file)

    with ResourceLock(file_path=input_file):

        with ResourceLock(file_path=output_file):
            with open(output_file, 'w') as outputStream:
                with open(input_file, 'r') as inputStream:
                    for line in inputStream:
                        outputStream.write('echo ' + escape_for_batch(line))


def escape_for_batch(text) -> str:
    special_chars: set[str] = {'%', '!', '^', '<', '>', '&', '|', '=', '+', ';', ',', '(', ')', '[', ']', '{', '}', '"'}
    new_line: str = ''
    # Escape each special character with a caret (^)
    for char in text:
        if char in special_chars:
            new_line += f'^{char}'
        else:
            new_line += char

    return new_line


def join_set_of_elements(container: set[str], delimiter: str) -> str:
    joined_set_str: str = ''
    for element in container:
        joined_set_str += element + delimiter
    return joined_set_str


def get_row_index_from_excel_cell_format(provided_cell_position: str):
    the_found_cell: Match[str] = re.fullmatch('^[A-Z]+\d+$', provided_cell_position)
    if the_found_cell is None:
        raise Exception("Invalid Excel cell format !")

    row_digits = []
    index: int = 0
    while index < provided_cell_position.__len__():
        current_char = provided_cell_position[index]
        if current_char.isdigit():
            row_digits.append(current_char)
        index += 1

    the_row_index = 0
    power_count = 0
    while row_digits.__len__() > 0:
        the_row_index = the_row_index + int(row_digits.pop()) * pow(10, power_count)
        power_count += 1

    return the_row_index


def extract_row_col_from_cell_pos_format(start_cell: str) -> Tuple[str, int]:
    result = re.search(r'([a-zA-Z]+)(\d+)', start_cell)
    if result:
        column = result.group(1)
        row = int(result.group(2))
        return column, row
    else:
        raise ValueError("Input does not match Excel cell position format")
