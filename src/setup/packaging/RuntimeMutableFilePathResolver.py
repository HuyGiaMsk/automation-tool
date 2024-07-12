import os
import sys

from src.setup.packaging.PathResolver import PathResolver


class RuntimeMutableFilePathResolver(PathResolver):
    __instance = None

    @staticmethod
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
            env_path_resolver_path: str = os.path.abspath(__file__)
            packing_dir: str = os.path.dirname(env_path_resolver_path)
            setup_dir: str = os.path.dirname(packing_dir)
            src_dir: str = os.path.dirname(setup_dir)
            root_repo_dir = os.path.dirname(src_dir)
            return root_repo_dir

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
            open(final_path).close()
            return final_path

        raise Exception(
            f'Something went wrong with {final_path} - No file/dir exist and could not create a new file/dir for it')
