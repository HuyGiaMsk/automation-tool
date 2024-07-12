import os
import sys

from src.common.RestrictCallers import restrict_callers
from src.setup.packaging.path.PathResolver import PathResolver
from src.setup.packaging.path.PathResolvingService import PathResolvingService


class InternalImmutableFilePathResolver(PathResolver):
    __instance = None

    @staticmethod
    @restrict_callers(PathResolvingService)
    def get_instance() -> PathResolver:
        if InternalImmutableFilePathResolver.__instance is None:
            InternalImmutableFilePathResolver.__instance = InternalImmutableFilePathResolver()
        return InternalImmutableFilePathResolver.__instance

    def __init__(self):
        # regarding the user files like log, input
        self.root_dir = self.get_executable_directory()

    def get_executable_directory(self) -> str:
        try:
            # inside the packaged executable file of Pyinstaller
            # _MEIPASS will be set at runtime, just discard the warning
            return sys._MEIPASS

        except Exception:
            # This is for running in an IDE or standard Python interpreter
            env_path_resolver_path: str = os.path.abspath(__file__)
            path_dir: str = os.path.dirname(env_path_resolver_path)
            packing_dir: str = os.path.dirname(path_dir)
            setup_dir: str = os.path.dirname(packing_dir)
            src_dir: str = os.path.dirname(setup_dir)
            root_repo_dir = os.path.dirname(src_dir)
            return root_repo_dir

    @restrict_callers(PathResolvingService)
    def resolve(self, paths: list[str]) -> str:

        if paths.__len__() == 0:
            return self.root_dir

        final_path: str = self.root_dir
        for path in paths:
            final_path = os.path.join(final_path, path)

        if not os.path.exists(final_path):
            raise Exception(
                f'Something went wrong, the expected path {final_path} is not existed in internal immutable paths')

        return final_path
