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
            root_dir: str = os.path.abspath(__file__)
            while not root_dir.endswith('automation-tool') or root_dir.endswith('automation_tool'):
                root_dir = os.path.dirname(root_dir)
                root_dir = root_dir.lower()
            return root_dir

    @only_accept_callers_from(PathResolvingService)
    def resolve(self, paths: list[str]) -> str:

        if paths.__len__() == 0:
            return self.root_dir

        final_path: str = self.root_dir
        for path in paths:
            final_path = os.path.join(final_path, path)

        if os.path.exists(final_path):
            return final_path

        if final_path.__contains__('.'):
            with open(final_path, 'w'):
                pass  # File created, do nothing
            return final_path

        os.mkdir(final_path)
        return final_path
