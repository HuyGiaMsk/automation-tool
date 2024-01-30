import os.path
from logging import Logger

from src.common.Constants import ROOT_DIR
from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger


def validate_keys_of_dictionary(settings: dict[str, str],
                                mandatory_settings: set[str]) -> None:
    logger: Logger = get_current_logger()
    error_messages: list[str] = []
    for key in mandatory_settings:

        value = settings.get(str(key))
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
                                  output_file: str = os.path.join(ROOT_DIR, 'output',
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
