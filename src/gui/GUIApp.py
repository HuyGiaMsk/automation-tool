import tkinter as tk
from tkinter import Label, Frame, Text, HORIZONTAL, ttk, messagebox
from tkinter.ttk import Combobox
import os
import importlib
from logging import Logger
from types import ModuleType

from src.gui.TextBoxLoggingHandler import setup_textbox_logger
from src.observer.Event import Event
from src.observer.EventBroker import EventBroker
from src.observer.EventHandler import EventHandler
from src.observer.PercentChangedEvent import PercentChangedEvent
from src.task.AutomatedTask import AutomatedTask
from src.common.Constants import ROOT_DIR
from src.common.FileUtil import load_key_value_from_file_properties
from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger


class GUIApp(tk.Tk, EventHandler):

    def __init__(self):
        super().__init__()
        self.automated_task = None

        self.protocol("WM_DELETE_WINDOW", self.handle_close_app)
        EventBroker.get_instance().subscribe(topic=PercentChangedEvent.event_name,
                                             observer=self)
        self.title("Automation Tool")
        self.geometry('1920x1080')

        self.logger: Logger = get_current_logger()

        self.container_frame = tk.Frame(self)
        self.container_frame.pack()

        self.myLabel = Label(self.container_frame, text='Automation Tool', font=('Maersk Headline Bold', 16))
        self.myLabel.pack()

        self.automated_tasks_dropdown = Combobox(
            master=self.container_frame,
            state="readonly",
        )
        self.automated_tasks_dropdown.pack()

        self.content_frame = Frame(self.container_frame, width=500, height=300, bd=1, relief=tk.SOLID)
        self.content_frame.pack(padx=20, pady=20)

        self.automated_tasks_dropdown.bind("<<ComboboxSelected>>", self.handle_task_dropdown_change)

        self.populate_task_dropdown()

        self.current_input_setting_values = {}
        self.current_automated_task_name = None

        self.custom_progressbar_text_style = ttk.Style()
        self.custom_progressbar_text_style.layout("Text.Horizontal.TProgressbar",
                                                  [('Horizontal.Progressbar.trough',
                                                    {'children': [('Horizontal.Progressbar.pbar',
                                                                   {'side': 'left', 'sticky': 'ns'}),
                                                                  ("Horizontal.Progressbar.label",
                                                                   {"sticky": ""})],
                                                     'sticky': 'nswe'})])
        self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar", text="None 0 %")

        self.progressbar = ttk.Progressbar(self.container_frame, orient=HORIZONTAL,
                                           length=500, mode="determinate", maximum=100
                                           , style="Text.Horizontal.TProgressbar")
        self.progressbar.pack(pady=20)

        self.is_task_currently_pause: bool = False
        self.pause_button = tk.Button(self.container_frame,
                                      text='Pause',
                                      command=lambda: self.handle_pause_button())
        self.pause_button.pack()

        self.terminate_button = tk.Button(self.container_frame,
                                          text='Terminate',
                                          command=lambda: self.handle_terminate_button())
        self.terminate_button.pack()

        self.textbox: Text = tk.Text(self.container_frame, wrap="word", state=tk.DISABLED, width=40, height=10)
        self.textbox.pack()
        setup_textbox_logger(self.textbox)

    def populate_task_dropdown(self):
        input_dir: str = os.path.join(ROOT_DIR, "input")
        automated_task_names: list[str] = []

        with ResourceLock(file_path=input_dir):
            for dir_name in os.listdir(input_dir):
                if dir_name.lower().endswith(".properties"):
                    automated_task_names.append(dir_name.replace(".properties", ""))

        automated_task_names.remove("InvokedClasses")
        self.automated_tasks_dropdown['values'] = automated_task_names

    def callback_before_run_task(self):
        setup_textbox_logger(self.textbox)

    def persist_settings_to_file(self):
        if self.current_automated_task_name is None:
            return

        file_path: str = os.path.join(ROOT_DIR, "input", "{}.properties".format(self.current_automated_task_name))
        self.logger.info("Attempt to persist data to {}".format(file_path))

        with ResourceLock(file_path=file_path):

            with open(file_path, 'w') as file:
                file.truncate(0)

            with open(file_path, 'a') as file:
                for key, value in self.current_input_setting_values.items():
                    file.write(f"{key} = {value}\n")

    def update_field_data(self, event):
        text_widget = event.widget
        new_value = text_widget.get("1.0", "end-1c")
        field_name = text_widget.special_id
        self.current_input_setting_values[field_name] = new_value
        self.logger.debug("Change data on field {} to {}".format(field_name, new_value))

    def update_frame_content(self, selected_task):
        # Clear the content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        # Create new content based on the selected task
        self.logger.info('Display fields for task {}'.format(selected_task))

        clazz_module: ModuleType = importlib.import_module('src.task.' + selected_task)
        clazz = getattr(clazz_module, selected_task)

        setting_file = os.path.join(ROOT_DIR, 'input', '{}.properties'.format(selected_task))
        input_setting_values: dict[str, str] = load_key_value_from_file_properties(setting_file)
        input_setting_values['invoked_class'] = selected_task
        input_setting_values['use.GUI'] = 'False'
        input_setting_values['time.unit.factor'] = '1'

        self.current_input_setting_values = input_setting_values
        self.current_automated_task_name = selected_task
        self.automated_task: AutomatedTask = clazz(input_setting_values, self.callback_before_run_task)

        for each_setting in input_setting_values:
            # Create a container frame for each label and text input pair
            setting_frame = Frame(self.content_frame)
            setting_frame.pack(anchor="w", pady=5)

            # Create the label and text input widgets inside the container frame
            field_label = Label(master=setting_frame, text=each_setting, width=15)
            field_label.pack(side="left")

            field_input = Text(master=setting_frame, width=30, height=1)
            field_input.pack(side="left")

            field_input.special_id = each_setting
            initial_value: str = input_setting_values.get(each_setting)
            field_input.insert("1.0", '' if initial_value is None else initial_value)
            field_input.bind("<KeyRelease>", self.update_field_data)

        perform_button = tk.Button(self.content_frame,
                                   text='Perform',
                                   font=('Maersk Headline Bold', 10),
                                   command=lambda: self.handle_click_on_perform_task_button(self.automated_task))
        perform_button.pack()

    def handle_close_app(self):
        self.persist_settings_to_file()
        self.destroy()

    def handle_incoming_event(self, event: Event) -> None:
        if isinstance(event, PercentChangedEvent):
            if self.automated_task is None:
                return

            current_task_name = type(self.automated_task).__name__
            if event.task_name is not current_task_name:
                return

            self.progressbar['value'] = event.current_percent
            self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar",
                                                         text="{} {}%".format(current_task_name,
                                                                              event.current_percent))

    def handle_task_dropdown_change(self, event):
        self.persist_settings_to_file()
        selected_task = self.automated_tasks_dropdown.get()
        self.update_frame_content(selected_task)

    def handle_click_on_perform_task_button(self, task: AutomatedTask):
        if task.is_alive():
            messagebox.showinfo("Have a task currently running",
                                "Please terminate the current task before run a new one")
            return

        self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar",
                                                     text="{} {}%".format(type(task).__name__, 0))
        task.start()

    def handle_pause_button(self):
        if self.automated_task.is_alive() is False:
            return

        if self.is_task_currently_pause:
            self.automated_task.resume()
            self.pause_button.config(text="Pause")
            self.is_task_currently_pause = False
            return

        self.automated_task.pause()
        self.pause_button.config(text="Resume")
        self.is_task_currently_pause = True
        return

    def handle_terminate_button(self):
        self.automated_task.terminate()
        self.automated_task = None
        self.progressbar['value'] = 0
        self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar",
                                                     text="{} {}%".format("None Task", 0))


if __name__ == "__main__":
    app = GUIApp()
    app.mainloop()
