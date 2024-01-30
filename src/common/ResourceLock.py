import time
import errno
import os

from src.common.ThreadLocalLogger import get_current_logger


class ResourceLockException(Exception):
    pass


class ResourceLock(object):
    """
        ResourceLock - a pessimistic lock mechanism based on file
        / support for multithreading
        / for running automated tasks simultaneously (in this project)
        Each task could deal with the same resource, I want when this task has privileges to a resource, it must keep it
        until complete the critical section of itself, avoid data inconsistency on files' content or folder files' existence

        A file locking mechanism that has context-manager support, so
        you can use it in a with statement. This should be relatively cross
        compatible as it doesn't rely on msvcrt or fcntl for the locking.
    """

    def __init__(self, file_path: str, timeout: int = 60 * 2, delay: float = 0.5,
                 log_make_clear_distinction_lock: str = None):
        """ Prepare the file locker. Specify the file to lock and optionally
            the maximum timeout and the delay between each attempt to lock.
        """
        self.__is_locked = False
        if os.path.isabs(file_path):
            self.lockfile = "%s.lock" % file_path
        else:
            self.lockfile = os.path.join(os.getcwd(), "%s.lock" % file_path)
        self.__file_path = file_path
        self.__timeout = timeout
        self.__delay = delay
        self.__open_lock_file = None
        self.__logged_content = '' if log_make_clear_distinction_lock is None else log_make_clear_distinction_lock

    def acquire(self):
        """ Acquire the lock, if possible. If the lock is in use, it checks again
            every `wait` seconds. It does this until it either gets the lock or
            exceeds `timeout` number of seconds, in which case it throws
            an exception.
        """
        logger = get_current_logger()
        start_time = time.time()
        while True:
            try:
                logger.debug("{} Try to lock {}".format("Append " + self.__logged_content, self.__file_path))
                self.__open_lock_file = os.open(self.lockfile, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                break

            except OSError as exception:
                if exception.errno != errno.EEXIST:
                    raise exception

                if (time.time() - start_time) >= self.__timeout:
                    logger.debug("Timeout occurred when try to get lock for file {}".format(self.__file_path))
                    raise ResourceLockException("Timeout occurred.")

                time.sleep(self.__delay)
        self.__is_locked = True

    def release(self):
        """ Get rid of the lock by deleting the lockfile.
            When working in a `with` statement, this gets automatically
            called at the end.
        """
        if self.__is_locked:
            os.close(self.__open_lock_file)
            os.unlink(self.lockfile)
            self.__is_locked = False

    def __enter__(self):
        """ Activated when used in the with statement.
            Should automatically acquire a lock to be used in the with block.
        """
        logger = get_current_logger()
        logger.debug("{} Enter __enter__".format(self.__logged_content))
        if not self.__is_locked:
            self.acquire()
        logger.debug("{} End __enter__".format(self.__logged_content))
        return self

    def __exit__(self, type, value, traceback):
        """ Activated at the end of the with statement.
            It automatically releases the lock if it isn't locked.
        """
        logger = get_current_logger()
        logger.debug("{} Enter __exit__".format(self.__logged_content))
        if type is not None:
            print(f"An exception of type {type} occurred with value {value}")

        if self.__is_locked:
            self.release()
        logger.debug("{} End __exit__".format(self.__logged_content))

    def __del__(self):
        """
            It is called when an object is about to be destroyed (garbage collected)
            Make sure that the FileLock instance doesn't leave a lockfile
            lying around.
        """
        self.release()
