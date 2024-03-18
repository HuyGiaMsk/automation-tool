import importlib
import os
import tkinter as tk
from logging import Logger
from tkinter import Label, Frame, Text, HORIZONTAL, ttk, messagebox, filedialog
from tkinter.ttk import Combobox
from types import ModuleType

from src.common.Constants import ROOT_DIR
from src.common.FileUtil import load_key_value_from_file_properties
from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger
from src.gui.TextBoxLoggingHandler import setup_textbox_logger
from src.observer.Event import Event
from src.observer.EventBroker import EventBroker
from src.observer.EventHandler import EventHandler
from src.observer.PercentChangedEvent import PercentChangedEvent
from src.task.AutomatedTask import AutomatedTask


class GUIApp(tk.Tk, EventHandler):

    def __init__(self):
        super().__init__()

        self.automated_task = None
        self.logger: Logger = get_current_logger()

        self.protocol("WM_DELETE_WINDOW", self.handle_close_app)
        EventBroker.get_instance().subscribe(topic=PercentChangedEvent.event_name,
                                             observer=self)
        self.title("Automation Tool")
        self.geometry('1080x980')
        self.configure(bg="#FFFFFF")

        self.container_frame = tk.Frame(self, bg="#FFFFFF")
        self.container_frame.pack()

        image_file = os.path.join(ROOT_DIR, "resource/img/logo3.png")
        self.logo_image = tk.PhotoImage(file=image_file)

        self.myLabel = Label(self.container_frame,
                             bg="#FFFFFF", width=980, image=self.logo_image,
                             compound=tk.CENTER)
        self.myLabel.pack()

        self.automated_tasks_dropdown = Combobox(master=self.container_frame, state="readonly", width=110, height=20,
                                                 background='#FB3D52', foreground='#FFFFFF')
        self.automated_tasks_dropdown.pack(padx=10, pady=10)

        self.content_frame = Frame(self.container_frame, width=1080, height=600, bd=1, relief=tk.SOLID, bg='#FFFFFF',
                                   borderwidth=0)
        self.content_frame.pack(padx=10, pady=10)

        self.automated_tasks_dropdown.bind("<<ComboboxSelected>>", self.handle_task_dropdown_change)
        self.populate_task_dropdown()

        self.current_input_setting_values = {}
        self.current_automated_task_name = None

        self.save_button = tk.Button(self.container_frame,
                                     text='Save',
                                     command=self.save_input,
                                     bg='#B678F2', fg='#FFFFFF', font=('Maersk Headline', 11),
                                     width=9, height=1, activeforeground='#FB3D52',
                                     )
        self.save_button.pack()

        self.is_task_currently_pause: bool = False
        self.pause_button = tk.Button(self.container_frame,
                                      text='Pause',
                                      command=lambda: self.handle_pause_button(),
                                      bg='#2FACE8', fg='#FFFFFF', font=('Maersk Headline', 11),
                                      width=9, height=1, activeforeground='#2FACE8'
                                      )
        self.pause_button.pack()

        self.terminate_button = tk.Button(self.container_frame,
                                          text='Reset',
                                          command=lambda: self.handle_terminate_button(),
                                          bg='#E34498', fg='#FFFFFF', font=('Maersk Headline', 11),
                                          width=9, height=1, activeforeground='#E34498'
                                          )
        self.terminate_button.pack()

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
        self.progressbar = ttk.Progressbar(self.container_frame, orient=HORIZONTAL,
                                           length=800, mode="determinate", maximum=100
                                           , style="Text.Horizontal.TProgressbar")
        self.progressbar.pack(pady=10)

        self.textbox: Text = tk.Text(self.container_frame, wrap="word", state=tk.DISABLED, width=100, height=15,
                                     background='#878787', font=('Maersk Text', 10), foreground='#FFFFFF')
        self.textbox.pack()
        setup_textbox_logger(self.textbox)

    def create_control_buttons(self):
        self.pause_button.pack(side=tk.LEFT, padx=5)
        self.terminate_button.pack(side=tk.LEFT, padx=5)
        self.save_button.pack(side=tk.LEFT, padx=5)

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
        self.logger.info("Data persisted successfully.")

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
        input_setting_values['time.unit.factor'] = '1'

        self.current_input_setting_values = input_setting_values
        self.current_automated_task_name = selected_task
        self.automated_task: AutomatedTask = clazz(input_setting_values, self.callback_before_run_task)

        for each_setting, initial_value in input_setting_values.items():
            if each_setting == 'use.GUI':
                continue

            # Create a container frame for each label and text input pair
            setting_frame = Frame(self.content_frame, background='#FFFFFF')
            setting_frame.pack(anchor="w", pady=5)

            # Create the label and text input widgets inside the container frame
            field_label = Label(master=setting_frame, text=each_setting, width=25,
                                font=('Maersk Text', 9), fg='#FFFFFF', bg='#FB3D52', borderwidth=0)
            field_label.pack(side="left")

            path_var = tk.StringVar()

            # Determine if the setting is a folder or file path
            if each_setting.endswith('.folder'):
                field_button = tk.Button(master=setting_frame, text="...",
                                         command=lambda var=path_var: self.choose_folder(var),
                                         height=1, borderwidth=0, bg='#FB3D52', fg='#FFFFFF')
            elif each_setting.endswith('.path'):
                field_button = tk.Button(master=setting_frame, text="...",
                                         command=lambda var=path_var: self.choose_file(var),
                                         height=1, borderwidth=0, bg='#FB3D52', fg='#FFFFFF')

            else:
                field_button = None

            if field_button:
                field_button.pack(side="right")

            field_input = Text(master=setting_frame, width=80, height=1, font=('Maersk Text', 9), background='#EDEDED',
                               fg='#000000', borderwidth=0)
            field_input.pack(side="left")

            field_input.special_id = each_setting
            field_input.insert("1.0", '' if initial_value is None else initial_value)

            field_input.bind("<KeyRelease>", self.update_field_data)

            # Cập nhật giá trị của StringVar vào field_input khi giá trị thay đổi
            path_var.trace("w", lambda *args, var=path_var, text=field_input: self.update_text_from_var(var, text))

        # Handle 'use.GUI' setting separately
        use_gui_var = tk.BooleanVar()
        use_gui_var.set(True if input_setting_values.get('use.GUI') == 'True' else False)
        use_gui_checkbox = tk.Checkbutton(self.content_frame, text="Use GUI", variable=use_gui_var,
                                          font=('Maersk Text', 9),
                                          background='#2FACE8', fg='#FFFFFF', width=21, height=1,
                                          )
        use_gui_checkbox.pack(anchor="w", pady=5)

        # Callback to update the setting value when the checkbox state changes
        use_gui_var.trace_add("write", lambda *args, var=use_gui_var: self.update_use_gui_setting(var))

        perform_button = tk.Button(self.content_frame,
                                   text='Perform',
                                   font=('Maersk Text', 11),
                                   command=lambda: self.handle_click_on_perform_task_button(self.automated_task),
                                   bg='#FB3D52', fg='#FFFFFF',
                                   width=9, height=1, activeforeground='#FB3D52')
        perform_button.pack(padx=5)

    def update_use_gui_setting(self, var):
        # Update the setting value based on the checkbox state
        self.current_input_setting_values['use.GUI'] = str(var.get())
        self.logger.debug("Change data on field use.GUI to {}".format(var.get()))

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
        self.update_save_button_state()

    def handle_click_on_perform_task_button(self, task: AutomatedTask):
        if task is not None and task.is_alive():
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

        self.automated_task = None
        self.progressbar['value'] = 0
        self.custom_progressbar_text_style.configure("Text.Horizontal.TProgressbar",
                                                     text="{} {}%".format("None Task", 0))

    # Function to handle choosing a folder
    def choose_folder(self, var):
        folder_selected = filedialog.askdirectory()

        if folder_selected:
            var.set(folder_selected)
            self.after(100, self.persist_settings_to_file)

    def choose_file(self, var):
        file_selected = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])

        if file_selected:
            var.set(file_selected)
            self.after(100, self.persist_settings_to_file)

    def update_text_from_var(self, var, text):
        text.delete("1.0", "end")
        text.insert("1.0", var.get())

    def check_all_fields_filled(self):
        for child in self.content_frame.winfo_children():
            if isinstance(child, Frame):
                for widget in child.winfo_children():
                    if isinstance(widget, Text):
                        if not widget.get("1.0", "end-1c").strip():  # Check if the field is empty
                            return False
        return True

    def update_save_button_state(self):
        if self.check_all_fields_filled():
            self.save_button.config(state=tk.NORMAL)
        else:
            self.save_button.config(state=tk.DISABLED)

    def save_input(self):
        """
        Create a dictionary to store the values of each field
        """

        saved_data = {}

        # Iterate over the fields to retrieve their values
        for child in self.content_frame.winfo_children():
            if isinstance(child, Frame):
                for widget in child.winfo_children():
                    if isinstance(widget, Text):
                        field_name = widget.special_id
                        field_value = widget.get("1.0", "end-1c").strip()
                        saved_data[field_name] = field_value

        # Determine the file name based on the selected task
        file_name = f"{self.current_automated_task_name}.properties"
        file_path = os.path.join(ROOT_DIR, "input", file_name)

        # Save the data to the file
        with open(file_path, 'w') as file:
            for field_name, field_value in saved_data.items():
                file.write(f"{field_name} = {field_value}\n")

        self.logger.info('Saved your input at {}'.format(file_path))


if __name__ == "__main__":
    app = GUIApp()
    app.mainloop()
