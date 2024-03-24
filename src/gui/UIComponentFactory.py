import tkinter as tk
from logging import Logger
from tkinter import filedialog, Text

from src.common.FileUtil import persist_settings_to_file
from src.common.ThreadLocalLogger import get_current_logger
from src.gui.UITaskPerformingStates import UITaskPerformingStates


class UIComponentFactory:
    _instance = None

    def __init__(self, app: UITaskPerformingStates):
        self.app = app

    @staticmethod
    def get_instance(app: UITaskPerformingStates):
        if app is None:
            raise Exception('Must provide the GUI app instance')

        if UIComponentFactory._instance is None:
            UIComponentFactory._instance = UIComponentFactory(app)
            return UIComponentFactory._instance

        if UIComponentFactory._instance.app is not app:
            UIComponentFactory._instance = UIComponentFactory(app)

        return UIComponentFactory._instance

    def create_component(self, setting_key: str, setting_value: str, parent_frame: tk.Frame) -> tk.Widget:
        setting_key_in_lowercase: str = setting_key.lower()

        if setting_key_in_lowercase.endswith('invoked_class'):
            return None

        if setting_key_in_lowercase.startswith('use.'):
            return self.create_checkbox_input(setting_key, setting_value, parent_frame)

        if setting_key_in_lowercase.endswith('.folder'):
            return self.create_folder_path_input(setting_key, setting_value, parent_frame)

        if setting_key_in_lowercase.endswith('.path'):
            return self.create_file_path_input(setting_key, setting_value, parent_frame)

        # default case, just text field
        return self.create_textbox_input(setting_key, setting_value, parent_frame)

    def create_textbox_input(self, setting_key: str, setting_value: str, parent_frame: tk.Frame) -> tk.Text:
        def update_field_data(event, app_states: UITaskPerformingStates):
            text_widget = event.widget
            new_value = text_widget.get("1.0", "end-1c")
            field_name = text_widget.special_id
            app_states.get_ui_settings()[field_name] = new_value
            app_states.get_task_instance().settings = app_states.get_ui_settings()
            persist_settings_to_file(app_states.get_task_name(), app_states.get_ui_settings())

            logger: Logger = get_current_logger()
            logger.debug("Change data on field {} to {}".format(field_name, new_value))

        field_label = tk.Label(master=parent_frame, text=setting_key, width=25,
                               font=('Maersk Text', 9), fg='#FFFFFF', bg='#FB3D52', borderwidth=0)
        field_label.pack(side="left")

        field_input = tk.Text(master=parent_frame, width=80, height=1, font=('Maersk Text', 9),
                              background='#EDEDED', fg='#000000', borderwidth=0)
        field_input.pack(side="left")
        field_input.special_id = setting_key
        field_input.insert("1.0", '' if setting_value is None else setting_value)
        field_input.bind("<KeyRelease>", lambda event, app=self.app: update_field_data(event, app))

        return field_input

    def create_folder_path_input(self, setting_key: str, setting_value: str, parent_frame: tk.Frame):

        def choosing_dir_callback(main_textbox: tk.Text):
            main_textbox.focus_set()
            dir_path = filedialog.askdirectory()
            if dir_path:
                main_textbox.delete("1.0", tk.END)
                main_textbox.insert(tk.END, dir_path)
                main_textbox.event_generate("<KeyRelease>")

        text_box: Text = self.create_textbox_input(setting_key, setting_value, parent_frame)
        btn_choose = tk.Button(master=parent_frame, text="Choose Folder",
                               command=lambda: choosing_dir_callback(text_box),
                               height=1, borderwidth=0, bg='#FB3D52', fg='#FFFFFF')
        btn_choose.pack(side="right")

        return text_box

    def create_file_path_input(self, setting_key: str, setting_value: str, parent_frame: tk.Frame):

        def choosing_file_callback(main_textbox: tk.Text):
            main_textbox.focus_set()
            file_path = filedialog.askopenfilename()
            if file_path:
                main_textbox.delete("1.0", tk.END)
                main_textbox.insert(tk.END, file_path)
                main_textbox.event_generate("<KeyRelease>")

        text_box: Text = self.create_textbox_input(setting_key, setting_value, parent_frame)
        btn_choose = tk.Button(master=parent_frame, text="Choose Folder",
                               command=lambda: choosing_file_callback(text_box),
                               height=1, borderwidth=0, bg='#FB3D52', fg='#FFFFFF')
        btn_choose.pack(side="right")

        return text_box

    def create_checkbox_input(self, setting_key: str, setting_value: str, parent_frame: tk.Frame) -> tk.Checkbutton:

        def updating_checkbox_callback(setting_name: str, value: tk.BooleanVar, app_states: UITaskPerformingStates):
            app_states.get_ui_settings()[setting_name] = str(value.get())
            app_states.get_task_instance().settings = app_states.get_ui_settings()
            persist_settings_to_file(app_states.get_task_name(), app_states.get_ui_settings())

        is_gui: bool = True if setting_value.lower() == 'true' else False

        is_checked = tk.BooleanVar(value=is_gui)

        use_gui_checkbox = tk.Checkbutton(parent_frame, text=setting_key,
                                          font=('Maersk Text', 9), background='#2FACE8',
                                          width=21, height=1,
                                          variable=is_checked,
                                          command=lambda: updating_checkbox_callback(setting_key, is_checked,
                                                                                     self.app))
        use_gui_checkbox.pack(anchor="w", pady=5)

        return use_gui_checkbox
