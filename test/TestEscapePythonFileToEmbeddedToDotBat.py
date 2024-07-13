import os

from src.common.StringUtil import escape_bat_file_special_chars
from src.setup.packaging.path.PathResolvingService import PathResolvingService


def escape_special_chars_to_embedded_python_to_bat():
    escape_bat_file_special_chars(
        input_file=os.path.join(PathResolvingService.get_instance().get_output_dir(), 'MinifiedDownloadSource.input'))


if __name__ == "__main__":
    escape_special_chars_to_embedded_python_to_bat()
