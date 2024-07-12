from src.setup.packaging.path.InternalImmutableFilePathResolver import InternalImmutableFilePathResolver
from src.setup.packaging.path.RuntimeMutableFilePathResolver import RuntimeMutableFilePathResolver


class PathResolvingService:
    __internal_immutable_dir_path: set[str] = {'src', 'test', 'resource'}

    @staticmethod
    def resolve(mandatory_path: str, *passing_paths) -> str:

        processing_paths: list[str] = [mandatory_path]
        is_tuple = isinstance(passing_paths, (list, tuple))
        if not is_tuple:
            raise TypeError("*paths must be a tuple of string elements")

        for path in passing_paths:
            if not isinstance(path, str):
                raise TypeError("*paths must be a string elements")
            processing_paths.append(path)

        if PathResolvingService.__internal_immutable_dir_path.__contains__(mandatory_path):
            return InternalImmutableFilePathResolver.get_instance().resolve(processing_paths)

        return RuntimeMutableFilePathResolver.get_instance().resolve(processing_paths)


TASK_DIR: str = PathResolvingService.resolve('src', 'task')
INPUT_DIR: str = PathResolvingService.resolve('input')
OUTPUT_DIR: str = PathResolvingService.resolve('output')
LOG_DIR: str = PathResolvingService.resolve('log')
