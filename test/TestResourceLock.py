import os
import threading
from logging import Logger

from src.common.Constants import ROOT_DIR
from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger

if __name__ == "__main__":
    path = os.path.join(ROOT_DIR, 'output', 'temp_file.txt')

    logger: Logger = get_current_logger()

    def write_lines(line, repeat=10):
        global logger  # Use the shared logger
        logger.info('{} Protecting file: {}'.format('Append ' + line[0], path))
        resource_lock = ResourceLock(path, log_make_clear_distinction_lock='Try to append ' + line[0])
        with resource_lock:
            for _ in range(repeat):
                with open(path, 'a') as tf:
                    tf.write(line + '\n')
                    tf.flush()


    th1 = threading.Thread(target=write_lines, args=('1111111111111111111111111111111', 10))
    th2 = threading.Thread(target=write_lines, args=('2222222222222222222222222222222', 10))
    th3 = threading.Thread(target=write_lines, args=('3333333333333333333333333333333', 10))
    th4 = threading.Thread(target=write_lines, args=('4444444444444444444444444444444', 10))

    th1.start()
    th2.start()
    th3.start()
    th4.start()

    th1.join()
    th2.join()
    th3.join()
    th4.join()

    assert not os.path.exists(f"%s.lock" % path), "The lock file wasn't cleaned up!"

    # Print the contents of the file.
    # Manually inspect the output.  Does it look like the operations were atomic?
    with open(path, "r") as f:
        print(f.read())

    os.unlink(path)
