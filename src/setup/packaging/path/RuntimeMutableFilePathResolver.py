import os
import sys

from src.common.RestrictCallers import only_accept_callers_from
from src.setup.packaging.path.PathResolver import PathResolver
from src.setup.packaging.path.PathResolvingService import PathResolvingService


class RuntimeMutableFilePathResolver(PathResolver):
    __instance = None

    @staticmethod
    @only_accept_callers_from(PathResolvingService)
    def get_instance() -> PathResolver:
        if RuntimeMutableFilePathResolver.__instance is None:
            RuntimeMutableFilePathResolver.__instance = RuntimeMutableFilePathResolver()
        return RuntimeMutableFilePathResolver.__instance

    def __init__(self):
        self.root_dir = self.get_executable_directory()

    def get_executable_directory(self) -> str:
        if getattr(sys, 'frozen', False):
            # This is set when running as an executable created by PyInstaller or similar tools
            exe_path: str = os.path.abspath(sys.executable)
            installation_dir: str = os.path.dirname(exe_path)
            return installation_dir
        else:
            # This is for running in an IDE or standard Python interpreter
            current_path: str = os.path.abspath(__file__)
            while not current_path.endswith('automation-tool') or current_path.endswith('automation_tool'):
                current_path = os.path.dirname(current_path)
                current_path = current_path.lower()
            return current_path

    @only_accept_callers_from(PathResolvingService)
    def resolve(self, paths: list[str]) -> str:

        if paths.__len__() == 0:
            return self.root_dir

        final_path: str = self.root_dir
        for path in paths:
            final_path = os.path.join(final_path, path)

        if os.path.exists(final_path):
            return final_path

        if os.path.isdir(final_path):
            os.mkdir(final_path)
            return final_path

        if os.path.isfile(final_path):
            with open(final_path, 'w'):
                pass  # File created, do nothing
            return final_path

        raise Exception(
            f'Something went wrong with {final_path} - No file/dir exist and could not create a new file/dir for it')
