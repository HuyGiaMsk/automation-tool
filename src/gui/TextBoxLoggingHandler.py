import logging
import tkinter as tk

from src.common.ThreadLocalLogger import get_current_logger


class TextBoxLoggingHandler(logging.Handler):

    def __init__(self, textbox):
        super().__init__()
        self.textbox = textbox

    def emit(self, record: str):
        msg = self.format(record)
        self.textbox.config(state=tk.NORMAL)
        self.textbox.insert(tk.END, msg + '\n')
        self.textbox.config(state=tk.DISABLED)
        self.textbox.see(tk.END)


def setup_textbox_logger(textbox: tk.Text):
    thread_local_logger: logging.Logger = get_current_logger()
    logging_handler: TextBoxLoggingHandler = TextBoxLoggingHandler(textbox)
    formatter: logging.Formatter = logging.Formatter(
        'GUI APP - %(asctime)s - %(levelname)s - %(filename)s %(funcName)s#%(lineno)d: %(message)s')
    logging_handler.setFormatter(formatter)
    thread_local_logger.addHandler(logging_handler)
