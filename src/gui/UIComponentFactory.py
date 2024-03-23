import tkinter as tk
from logging import Logger
from tkinter import filedialog, Text

from src.common.FileUtil import persist_settings_to_file
from src.common.ThreadLocalLogger import get_current_logger
from src.gui.GUIApp import GUIApp


class UIComponentFactory:
    _instance = None

    _app: GUIApp = None

    @staticmethod
    def get_instance(app: GUIApp):
        if UIComponentFactory._instance is None:
            UIComponentFactory._instance = UIComponentFactory()

        if app is None:
            raise Exception('Must provide the GUI app instance')

        _app = app
        return UIComponentFactory._instance

    def create_component(self, setting_key: str, setting_value: str, parent_frame: tk.Frame) -> tk.Widget:
        setting_key = setting_key.lower()

        if setting_key.endswith('.gui'):
            return self.create_checkbox(setting_key, setting_value, parent_frame)

        if setting_key.endswith('.folder'):
            return self.create_folder_path_input(setting_value, setting_value, parent_frame)

        if setting_key.endswith('.path'):
            return self.create_folder_path_input(setting_value, setting_value, parent_frame)

        return self.create_text_input_field(setting_value, setting_value, parent_frame)

    def create_text_input_field(self, setting_key: str, setting_value: str, parent_frame: tk.Frame) -> tk.Text:
        def update_field_data(event):
            text_widget = event.widget
            new_value = text_widget.get("1.0", "end-1c")
            field_name = text_widget.special_id
            self._app.current_input_setting_values[field_name] = new_value
            persist_settings_to_file(self._app.current_automated_task_name, self._app.current_input_setting_values)

            logger: Logger = get_current_logger()
            logger.debug("Change data on field {} to {}".format(field_name, new_value))

        field_input = tk.Text(master=parent_frame, width=80, height=1, font=('Maersk Text', 9),
                              background='#EDEDED', fg='#000000', borderwidth=0)
        field_input.pack(side="left")
        field_input.special_id = setting_key
        field_input.insert("1.0", '' if setting_value is None else setting_value)
        field_input.bind("<KeyRelease>", update_field_data)

        return field_input

    def create_folder_path_input(self, setting_key: str, setting_value: str, parent_frame: tk.Frame):

        def choose_folder(main_textbox: tk.Text):
            file_path = filedialog.askdirectory()
            if file_path:
                main_textbox.delete(0, tk.END)  # Clear any existing text
                main_textbox.insert(0, file_path)  # Insert the chosen file path

        text_box: Text = self.create_text_input_field(setting_key, setting_value, parent_frame)
        btn_choose = tk.Button(master=parent_frame, text="Choose Folder", command=lambda: choose_folder(text_box),
                               height=1, borderwidth=0, bg='#FB3D52', fg='#FFFFFF')
        btn_choose.pack(side="right")

        return text_box

    def create_file_path_input(self, setting_key: str, setting_value: str, parent_frame: tk.Frame):

        def choose_file(main_textbox: tk.Text):
            file_path = filedialog.askopenfilename()
            if file_path:
                main_textbox.delete(0, tk.END)  # Clear any existing text
                main_textbox.insert(0, file_path)  # Insert the chosen file path

        text_box: Text = self.create_text_input_field(setting_key, setting_value, parent_frame)
        btn_choose = tk.Button(master=parent_frame, text="Choose Folder", command=lambda: choose_file(text_box),
                               height=1, borderwidth=0, bg='#FB3D52', fg='#FFFFFF')
        btn_choose.pack(side="right")

        return text_box

    def create_checkbox(self, setting_key: str, setting_value: str, parent_frame: tk.Frame) -> tk.Checkbutton:

        def updating_use_gui_checkbox_callback(setting_name: str, check_button: tk.Checkbutton):
            self._app.current_input_setting_values[setting_name] = check_button['variable'].get()

            if self._app.content_frame is not None:
                persist_settings_to_file(self._app.current_automated_task_name, self._app.current_input_setting_values)

        is_gui: bool = True if setting_value.lower() == 'True'.lower() else False

        use_gui_checkbox = tk.Checkbutton(parent_frame,
                                          text="Use GUI", font=('Maersk Text', 9), background='#2FACE8',
                                          width=21, height=1,
                                          command=lambda: updating_use_gui_checkbox_callback(setting_key,
                                                                                             use_gui_checkbox))
        if is_gui:
            use_gui_checkbox.select()

        use_gui_checkbox.pack(anchor="w", pady=5)

        return use_gui_checkbox
