import os
import tkinter as tk
from logging import Logger
from tkinter import Label, Frame, Text, HORIZONTAL, ttk, messagebox, Button
from tkinter.ttk import Combobox, Progressbar, Style
from typing import Tuple

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

    def __init__(self):
        super().__init__()
        # Register this gui app instance as an observer listening for PercentChangedEvent <-> effect changes
        # in progress bar
        EventBroker.get_instance().subscribe(topic=PercentChangedEvent.event_name,
                                             observer=self)
        # logger for GUI thread
        self.logger: Logger = get_current_logger()

        # UI Task Performing States
        self.automated_task: AutomatedTask
        self.current_task_settings: dict[str, str] = {}
        self.current_task_name = None
        self.is_task_currently_pause: bool = False

        # basic configurations for the Tk instance
        self.title("Automation Tool")
        self.geometry('1080x980')
        self.configure(bg="#FFFFFF")

        # register the life cycle callback when before ending/closing the tk instance/window
        self.protocol("WM_DELETE_WINDOW", self.handle_close_app)

        # initial rendering - layout portions
        whole_app_frame = tk.Frame(self, bg="#FFFFFF")
        whole_app_frame.pack()

        self.logo_image: tk.PhotoImage = tk.PhotoImage(file=os.path.join(ROOT_DIR, "resource/img/logo5.png"))

        self.render_header(parent_frame=whole_app_frame, logo=self.logo_image)

        tasks_dropdown: Combobox = self.render_tasks_dropdown(parent_frame=whole_app_frame)

        self.main_content_frame = Frame(master=whole_app_frame, width=1080, height=600, bd=1, relief=tk.SOLID,
                                        bg='#FFFFFF', borderwidth=0)
        self.main_content_frame.pack(padx=10, pady=10)

        self.render_main_content_frame_for_first_task(tasks_dropdown=tasks_dropdown)

        self.progress_bar, self.progress_bar_label = self.render_progress_bar(parent_frame=whole_app_frame)

        self.logging_textbox = self.render_textbox_logger(parent_frame=whole_app_frame)

    def get_ui_settings(self) -> dict[str, str]:
        return self.current_task_settings

    def set_ui_settings(self, new_ui_setting_values: dict[str, str]) -> dict[str, str]:
        self.current_task_settings = new_ui_setting_values
        return self.current_task_settings

    def get_task_name(self) -> str:
        return self.current_task_name

    def get_task_instance(self) -> AutomatedTask:
        return self.automated_task

    # This ui app will act as an observer, listening/handling the event from the publisher
    def handle_incoming_event(self, event: Event) -> None:
        if isinstance(event, PercentChangedEvent):
            if self.automated_task is None:
                return

            current_task_name = type(self.automated_task).__name__
            if event.task_name is not current_task_name:
                return

            if event.current_percent - float(self.progress_bar['value']) > 10:
                return

            self.progress_bar['value'] = event.current_percent
            self.progress_bar_label.configure("Text.Horizontal.TProgressbar",
                                              text="{} {}%".format(current_task_name,
                                                                   event.current_percent))

    # Life cycle callback before closing the ui app
    def handle_close_app(self) -> None:
        persist_settings_to_file(self.current_task_name, self.current_task_settings)
        self.destroy()

    def render_header(self, parent_frame: Frame, logo: tk.PhotoImage) -> Label:
        logo_label: Label = Label(parent_frame, bg="#FFFFFF", width=980, image=logo, compound=tk.CENTER)
        logo_label.pack()
        return logo_label

    def render_tasks_dropdown(self, parent_frame: Frame) -> Combobox:
        tasks_dropdown: Combobox = Combobox(master=parent_frame, state="readonly", width=110, height=20,
                                            background='#FB3D52', foreground='#FFFFFF')
        tasks_dropdown.pack(padx=10, pady=10)
        tasks_dropdown.bind("<<ComboboxSelected>>", self.handle_tasks_dropdown)
        self.populate_task_dropdown(tasks_dropdown)
        return tasks_dropdown

    # Find all available defined tasks and populate these as values of a dropdown
    def populate_task_dropdown(self, dropdown: Combobox):
        input_dir: str = os.path.join(ROOT_DIR, 'src', 'task')
        automated_task_names: list[str] = []

        with ResourceLock(file_path=input_dir):
            for dir_name in os.listdir(input_dir):
                if dir_name.lower().endswith(".py"):
                    clean_name = dir_name.replace(".py", "")
                    automated_task_names.append(clean_name)

        automated_task_names.remove("AutomatedTask")
        automated_task_names.remove("__init__")
        dropdown['values'] = automated_task_names

    def handle_tasks_dropdown(self, event):
        if self.current_task_name is not None and self.current_task_settings is not None:
            persist_settings_to_file(self.current_task_name, self.current_task_settings)

        selected_task = event.widget.get()
        self.render_main_content_frame(selected_task)

    def render_main_content_frame(self, selected_task):
        # Clear the content frame
        for widget in self.main_content_frame.winfo_children():
            widget.destroy()

        # Create new content based on the selected task
        self.logger.info('Display fields for task {}'.format(selected_task))

        setting_file = os.path.join(ROOT_DIR, 'input', '{}.properties'.format(selected_task))
        if not os.path.exists(setting_file):
            with open(setting_file, 'w'):
                pass  # File created, do nothing

        input_setting_values: dict[str, str] = load_key_value_from_file_properties(setting_file)
        input_setting_values['invoked_class'] = selected_task

        if input_setting_values.get('time.unit.factor') is None:
            input_setting_values['time.unit.factor'] = '1'

        if input_setting_values.get('use.GUI') is None:
            input_setting_values['use.GUI'] = 'True'

        self.automated_task = create_task_instance(input_setting_values,
                                                   selected_task,
                                                   lambda: setup_textbox_logger(self.logging_textbox))
        mandatory_settings: list[str] = self.automated_task.mandatory_settings()
        mandatory_settings.append('invoked_class')
        mandatory_settings.append('time.unit.factor')
        mandatory_settings.append('use.GUI')

        self.current_task_name = selected_task
        self.current_task_settings = {}
        for each_setting in mandatory_settings:
            # Create a container frame for each pair combining a label and an input
            setting_frame = Frame(self.main_content_frame, background='#FFFFFF')
            setting_frame.pack(anchor="w", pady=5)

            initial_value: str = input_setting_values.get(each_setting)
            self.current_task_settings[each_setting] = initial_value
            UIComponentFactory.get_instance(self).create_component(each_setting, initial_value, setting_frame)

        self.automated_task.settings = self.current_task_settings
        self.render_button_frame()

    def render_button_frame(self):
        button_frame = tk.Frame(master=self.main_content_frame, bg='#FFFFFF')
        button_frame.pack(expand=True, fill="both")
        # Create a left and right frame with a flexible column configuration
        left_frame = tk.Frame(master=button_frame, bg='#FFFFFF')
        left_frame.pack(side="left", expand=True, fill="both")
        right_frame = tk.Frame(master=button_frame, bg='#FFFFFF')
        right_frame.pack(side="right", expand=True, fill="both")
        perform_button = tk.Button(button_frame,
                                   text='Perform',
                                   font=('Maersk Text', 11),
                                   command=lambda: self.handle_perform_button(),
                                   bg='#2FACE8', fg='#FFFFFF',
                                   width=9, height=1, activeforeground='#FB3D52')
        perform_button.pack(side='left')

        self.pause_button = self.render_pause_button(parent_frame=button_frame)
        self.render_reset_button(parent_frame=button_frame)

    def render_main_content_frame_for_first_task(self, tasks_dropdown: Combobox):
        tasks_dropdown.focus_set()
        tasks_dropdown.current(0)
        tasks_dropdown.event_generate("<<ComboboxSelected>>")

    def handle_perform_button(self):
        if self.automated_task is not None and self.automated_task.is_alive():
            messagebox.showinfo("Have a task currently running",
                                "Please terminate the current task before run a new one")
            return

        if self.automated_task is None:
            self.automated_task = create_task_instance(self.current_task_settings,
                                                       self.current_task_name,
                                                       lambda: setup_textbox_logger(self.logging_textbox))

        self.progress_bar_label.configure("Text.Horizontal.TProgressbar",
                                          text="{} {}%".format(type(self.automated_task).__name__, 0))
        self.automated_task.start()

    def render_pause_button(self, parent_frame: Frame) -> tk.Button:
        pause_button: Button = tk.Button(master=parent_frame, text='Pause', command=self.handle_pause_button,
                                         bg='#FA6A55', fg='#FFFFFF', font=('Maersk Headline', 11), width=9, height=1,
                                         activeforeground='#FA6A55')
        pause_button.pack(side='left')
        return pause_button

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

    def render_reset_button(self, parent_frame: Frame) -> Button:
        reset_button: Button = tk.Button(parent_frame, text='Reset', command=self.handle_reset_button,
                                         bg='#00243D', fg='#FFFFFF', font=('Maersk Headline', 11), width=9, height=1,
                                         activeforeground='#00243D')
        reset_button.pack(side='left')
        return reset_button

    def handle_reset_button(self):
        if self.automated_task:
            self.automated_task.terminate()

        if self.is_task_currently_pause:
            self.pause_button.config(text="Pause")
            self.is_task_currently_pause = False

        self.automated_task = None
        if self.automated_task is None:
            self.automated_task = create_task_instance(self.current_task_settings,
                                                       self.current_task_name,
                                                       lambda: setup_textbox_logger(self.logging_textbox))
        self.progress_bar['value'] = 0
        self.progress_bar_label.configure("Text.Horizontal.TProgressbar",
                                          text="{} {}%".format("None Task", 0))

    def render_progress_bar(self, parent_frame: Frame) -> Tuple[Progressbar, Style]:
        progressbar_text = ttk.Style()
        progressbar_text.layout("Text.Horizontal.TProgressbar",
                                [('Horizontal.Progressbar.trough', {
                                    'children': [('Horizontal.Progressbar.pbar', {'side': 'left', 'sticky': 'ns'}),
                                                 ("Horizontal.Progressbar.label", {"sticky": ""})],
                                    'sticky': 'nswe'})])
        progressbar_text.configure(style="Text.Horizontal.TProgressbar", text="None 0 %", background='#FB3D52',
                                   troughcolor='#40AB35', troughrelief='flat', bordercolor='#FB3D52',
                                   lightcolor='#42B0D5', darkcolor='#00243D')

        progressbar: Progressbar = ttk.Progressbar(parent_frame, orient=HORIZONTAL, length=800, mode="determinate",
                                                   maximum=100, style="Text.Horizontal.TProgressbar")
        progressbar.pack(pady=10)
        return progressbar, progressbar_text

    def render_textbox_logger(self, parent_frame: Frame):
        textbox: Text = tk.Text(master=parent_frame, wrap="word", state=tk.DISABLED, width=100, height=15,
                                background='#878787', font=('Maersk Text', 10), foreground='#FFFFFF')
        textbox.pack()
        setup_textbox_logger(textbox)
        return textbox


if __name__ == "__main__":
    app = GUIApp()
    app.mainloop()
