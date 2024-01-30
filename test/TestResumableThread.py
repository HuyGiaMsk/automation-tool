import time
from src.common.ResumableThread import ResumableThread


class Test(ResumableThread):
    def perform(self):

        for i in range(100):

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

                print(f"Task running: {i}")
                time.sleep(1)


# Create a resumable thread with the example_task as its target
resumable_thread = Test()

# Start the thread
resumable_thread.start()

# Let the thread run for a while
time.sleep(5)

# Pause the thread
print("Pausing the thread")
resumable_thread.pause()

# Let the thread pause for a while
time.sleep(10)

# Resume the thread
print("Resuming the thread")
resumable_thread.resume()

# Let the thread run for a while
time.sleep(2)

# Terminate the thread
print("Terminating the thread")
resumable_thread.terminate()

# # Wait for the thread to finish
# resumable_thread.join()
#
# print("Thread terminated")