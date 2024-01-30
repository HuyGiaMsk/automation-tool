import os

from src.common.Constants import ROOT_DIR
from src.common.StringUtil import escape_bat_file_special_chars


def escape_special_chars_to_embedded_python_to_bat():
    escape_bat_file_special_chars(input_file=os.path.join(ROOT_DIR, 'input', 'MinifiedDownloadSource.input'))


if __name__ == "__main__":
    escape_special_chars_to_embedded_python_to_bat()
