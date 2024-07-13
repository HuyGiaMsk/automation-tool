import os
import threading
from logging import Logger
from threading import Thread

from src.common.FileUtil import load_key_value_from_file_properties
from src.common.ReflectionUtil import create_task_instance
from src.common.StringUtil import validate_keys_of_dictionary
from src.common.ThreadLocalLogger import get_current_logger
from src.setup.packaging.path.PathResolvingService import PathResolvingService
from src.task.AutomatedTask import AutomatedTask

if __name__ == "__main__":
    input_dir: str = PathResolvingService.get_instance().get_input_dir()
    setting_file: str = os.path.join(input_dir, 'InvokedClasses.properties')
    if not os.path.exists(setting_file):
        raise Exception("The settings file {} is not existed. Please providing it !".format(setting_file))

    settings: dict[str, str] = load_key_value_from_file_properties(setting_file)
    validate_keys_of_dictionary(settings, {'invoked_classes', 'run.sequentially'})
    defined_classes: list[str] = [class_name.strip() for class_name in settings['invoked_classes'].split(',')]
    run_sequentially: bool = 'True'.lower() == str(settings['run.sequentially']).lower()

    running_threads: list[Thread] = []
    for invoked_class in defined_classes:

        logger: Logger = get_current_logger()
        logger.info('Invoking class {}'.format(invoked_class))

        setting_file = os.path.join(input_dir, '{}.properties'.format(invoked_class))
        if not os.path.exists(setting_file):
            raise Exception("The settings file {} is not existed. Please providing it !".format(setting_file))
        
        settings: dict[str, str] = load_key_value_from_file_properties(setting_file)
        settings['invoked_class'] = invoked_class
        automated_task: AutomatedTask = create_task_instance(settings, invoked_class, None)

        if run_sequentially:
            automated_task.perform()
            continue

        # run concurrently
        running_task_thread: Thread = threading.Thread(target=automated_task.perform,
                                                       daemon=False)
        running_task_thread.start()
        running_threads.append(running_task_thread)

    for thread in running_threads:
        thread.join(timeout=60 * 60)
