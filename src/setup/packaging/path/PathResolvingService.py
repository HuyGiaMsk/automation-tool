class PathResolvingService:
    __instance = None

    @staticmethod
    def get_instance():
        if PathResolvingService.__instance is None:
            PathResolvingService.__instance = PathResolvingService()
        return PathResolvingService.__instance

    def __init__(self):
        self.__internal_immutable_dir_path: set[str] = {'src', 'test', 'resource'}
        self.__task_dir: str = None
        self.__input_dir: str = None
        self.__output_dir: str = None
        self.__log_dir: str = None

    def resolve(self, mandatory_path: str, *passing_paths) -> str:

        processing_paths: list[str] = [mandatory_path]
        is_tuple = isinstance(passing_paths, (list, tuple))
        if not is_tuple:
            raise TypeError("*paths must be a list/tuple of string elements")

        for path in passing_paths:
            if not isinstance(path, str):
                raise TypeError("*paths must be a string elements")
            processing_paths.append(path)

        if self.__internal_immutable_dir_path.__contains__(mandatory_path):
            from src.setup.packaging.path.InternalImmutableFilePathResolver import InternalImmutableFilePathResolver
            return InternalImmutableFilePathResolver.get_instance().resolve(processing_paths)

        from src.setup.packaging.path.RuntimeMutableFilePathResolver import RuntimeMutableFilePathResolver
        return RuntimeMutableFilePathResolver.get_instance().resolve(processing_paths)

    def get_task_dir(self):
        if self.__task_dir is None:
            self.__task_dir = self.resolve('src', 'task')
        return self.__task_dir

    def get_input_dir(self):
        if self.__input_dir is None:
            self.__input_dir = self.resolve('input')
        return self.__input_dir

    def get_output_dir(self):
        if self.__output_dir is None:
            self.__output_dir = self.resolve('output')
        return self.__output_dir

    def get_log_dir(self):
        if self.__log_dir is None:
            self.__log_dir = self.resolve('log')
        return self.__log_dir
