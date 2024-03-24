import os
import tkinter as tk
from logging import Logger
from tkinter import Label, Frame, Text, HORIZONTAL, ttk, messagebox, Widget
from tkinter.ttk import Combobox

from src.common.Constants import ROOT_DIR
from src.common.FileUtil import load_key_value_from_file_properties, persist_settings_to_file
from src.common.ReflectionUtil import create_task_instance
from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger
from src.gui.TextBoxLoggingHandler import setup_textbox_logger
from src.gui.UIComponentFactory import UIComponentFactory
from src.gui.UITaskPerformingStates import UITaskPerformingStates
from src.observer.Event import Event
from src.observer.EventBroker import EventBroker
from src.observer.EventHandler import EventHandler
from src.observer.PercentChangedEvent import PercentChangedEvent
from src.task.AutomatedTask import AutomatedTask


class GUIApp(tk.Tk, EventHandler, UITaskPerformingStates):

    def get_ui_settings(self) -> dict[str, str]:
        return self.current_task_settings

    def set_ui_settings(self, new_ui_setting_values: dict[str, str]) -> dict[str, str]:
        self.current_task_settings = new_ui_setting_values
        return self.current_task_settings

    def get_task_name(self) -> str:
        return self.current_task_name

    def get_task_instance(self) -> AutomatedTask:
        return self.automated_task

    def __init__(self):
        super().__init__()

        self.logger: Logger = get_current_logger()
        self.automated_task: AutomatedTask = None
        self.current_task_settings: dict[str, str] = {}
        self.current_task_name = None

        self.protocol("WM_DELETE_WINDOW", self.handle_close_app)
        EventBroker.get_instance().subscribe(topic=PercentChangedEvent.event_name,
                                             observer=self)
        self.title("Automation Tool")
        self.geometry('1080x980')
        self.configure(bg="#FFFFFF")

        self.whole_app_frame = tk.Frame(self, bg="#FFFFFF")
        self.whole_app_frame.pack()

        self.logo_image = tk.PhotoImage(file=os.path.join(ROOT_DIR, "resource/img/logo3.png"))

        self.myLabel = Label(self.whole_app_frame,
                             bg="#FFFFFF", width=980, image=self.logo_image,
                             compound=tk.CENTER)
        self.myLabel.pack()

        self.automated_tasks_dropdown = Combobox(master=self.whole_app_frame, state="readonly", width=110, height=20,
                                                 background='#FB3D52', foreground='#FFFFFF')
        self.automated_tasks_dropdown.pack(padx=10, pady=10)

        self.main_content_frame = Frame(self.whole_app_frame, width=1080, height=600, bd=1, relief=tk.SOLID,
                                        bg='#FFFFFF',
                                        borderwidth=0)
        self.main_content_frame.pack(padx=10, pady=10)

        self.automated_tasks_dropdown.bind("<<ComboboxSelected>>", self.handle_tasks_dropdown)
        self.populate_task_dropdown()

        self.is_task_currently_pause: bool = False
        self.pause_button = tk.Button(self.whole_app_frame,
                                      text='Pause',
                                      command=lambda: self.handle_pause_button(),
                                      bg='#2FACE8', fg='#FFFFFF', font=('Maersk Headline', 11),
                                      width=9, height=1, activeforeground='#2FACE8'
                                      )
        self.pause_button.pack()

        self.reset_button = tk.Button(self.whole_app_frame,
                                      text='Reset',
                                      command=lambda: self.handle_reset_button(),
                                      bg='#E34498', fg='#FFFFFF', font=('Maersk Headline', 11),
                                      width=9, height=1, activeforeground='#E34498'
                                      )
        self.reset_button.pack()

        self.custom_progressbar_text_style = ttk.Style()
        self.custom_progressbar_text_style.layout("Text.Horizontal.TProgressbar",
                                                  [('Horizontal.Progressbar.trough',
                                                    {'children': [('Horizontal.Progressbar.pbar',
                                                                   {'side': 'left', 'sticky': 'ns'}),
                                                                  ("Horizontal.Progressbar.label",
                                                                   {"sticky": ""})],
                                                     'sticky': 'nswe'})])
        self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar", text="None 0 %",
                                                     background='#FB3D52', troughcolor='#FB3D52',
                                                     troughrelief='flat', bordercolor='#FB3D52',
                                                     lightcolor='#FB3D52', darkcolor='#FB3D52'
                                                     )
        self.progressbar = ttk.Progressbar(self.whole_app_frame, orient=HORIZONTAL,
                                           length=800, mode="determinate", maximum=100
                                           , style="Text.Horizontal.TProgressbar")
        self.progressbar.pack(pady=10)

        self.textbox: Text = tk.Text(self.whole_app_frame, wrap="word", state=tk.DISABLED, width=100, height=15,
                                     background='#878787', font=('Maersk Text', 10), foreground='#FFFFFF')
        self.textbox.pack()
        setup_textbox_logger(self.textbox)

    # Life cycle callback before closing the ui app
    def handle_close_app(self):
        persist_settings_to_file(self.current_task_name, self.current_task_settings)
        self.destroy()

    # This ui app will act as an observer, listening/handling the event from the publisher
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

    # Find all available defined tasks and populate these as values of a dropdown
    def populate_task_dropdown(self):
        input_dir: str = os.path.join(ROOT_DIR, 'src', 'task')
        automated_task_names: list[str] = []

        with ResourceLock(file_path=input_dir):
            for dir_name in os.listdir(input_dir):
                if dir_name.lower().endswith(".py"):
                    clean_name = dir_name.replace(".py", "")
                    automated_task_names.append(clean_name)

        automated_task_names.remove("AutomatedTask")
        automated_task_names.remove("__init__")
        self.automated_tasks_dropdown['values'] = automated_task_names

    def render_main_content_frame(self, selected_task):
        # Clear the content frame
        for widget in self.main_content_frame.winfo_children():
            widget.destroy()

        # Create new content based on the selected task
        self.logger.info('Display fields for task {}'.format(selected_task))

        setting_file = os.path.join(ROOT_DIR, 'input', '{}.properties'.format(selected_task))
        input_setting_values: dict[str, str] = load_key_value_from_file_properties(setting_file)
        input_setting_values['invoked_class'] = selected_task
        input_setting_values['time.unit.factor'] = '1'
        if input_setting_values.get('use.GUI') is None:
            input_setting_values['use.GUI'] = 'True'

        self.automated_task = create_task_instance(input_setting_values,
                                                   selected_task,
                                                   lambda: setup_textbox_logger(self.textbox))
        self.current_task_name = selected_task
        self.current_task_settings = {}
        mandatory_settings: list[str] = self.automated_task.mandatory_settings()
        mandatory_settings.append('invoked_class')
        mandatory_settings.append('time.unit.factor')
        mandatory_settings.append('use.GUI')

        for each_setting in mandatory_settings:
            # Create a container frame for each pair combining a label and an input
            setting_frame = Frame(self.main_content_frame, background='#FFFFFF')
            setting_frame.pack(anchor="w", pady=5)

            initial_value: str = input_setting_values.get(each_setting)
            self.current_task_settings[each_setting] = initial_value
            child_component: Widget = UIComponentFactory.get_instance(self).create_component(
                each_setting, initial_value, setting_frame)

        self.automated_task.settings = self.current_task_settings
        perform_button = tk.Button(self.main_content_frame,
                                   text='Perform',
                                   font=('Maersk Text', 11),
                                   command=lambda: self.handle_perform_button(),
                                   bg='#FB3D52', fg='#FFFFFF',
                                   width=9, height=1, activeforeground='#FB3D52')
        perform_button.pack(padx=5)

    def handle_tasks_dropdown(self, event):
        persist_settings_to_file(self.current_task_name, self.current_task_settings)
        selected_task = self.automated_tasks_dropdown.get()
        self.render_main_content_frame(selected_task)

    def handle_perform_button(self):
        if self.automated_task is not None and self.automated_task.is_alive():
            messagebox.showinfo("Have a task currently running",
                                "Please terminate the current task before run a new one")
            return

        if self.automated_task is None:
            self.automated_task = create_task_instance(self.current_task_settings,
                                                       self.current_task_name,
                                                       lambda: setup_textbox_logger(self.textbox))

        self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar",
                                                     text="{} {}%".format(type(self.automated_task).__name__, 0))
        self.automated_task.start()

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

    def handle_reset_button(self):
        if self.automated_task:
            self.automated_task.terminate()

        if self.is_task_currently_pause:
            self.pause_button.config(text="Pause")
            self.is_task_currently_pause = False

        self.automated_task = None
        self.progressbar['value'] = 0
        self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar",
                                                     text="{} {}%".format("None Task", 0))


if __name__ == "__main__":
    app = GUIApp()
    app.mainloop()
