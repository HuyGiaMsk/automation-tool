import inspect
from types import FrameType
from typing import Callable, Union, Tuple


def restrict_callers(*optional_callers: Union[type, Tuple[type, ...]]) -> Callable:
    def decorator(func: Callable) -> Callable:
        def wrapper(*args, **kwargs):
            # Set of accepted caller classes
            accepted_callers = set(optional_callers)

            previous_frame: FrameType = inspect.currentframe().f_back

            caller_locals = previous_frame.f_locals
            caller_globals = previous_frame.f_globals

            # Check if 'self' or 'cls' is in locals and is an instance of an accepted caller
            if 'self' in caller_locals and isinstance(caller_locals['self'], tuple(accepted_callers)):
                return func(*args, **kwargs)
            if 'cls' in caller_locals and isinstance(caller_locals['cls'], tuple(accepted_callers)):
                return func(*args, **kwargs)

            # Check if we're in a static method context
            if '__name__' in caller_globals:
                # Extract the module and class name from the globals
                module_name = caller_globals['__name__']
                for cls in accepted_callers:
                    if cls.__module__ == module_name and cls.__name__ in caller_globals:
                        return func(*args, **kwargs)

            raise Exception(f'We should not call {func} directly ! Please use via the PathResolvingService !')

        return wrapper

    return decorator
