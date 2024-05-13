import os
from typing import Callable

import xlwings as xw

from src.task.AutomatedTask import AutomatedTask


class Convert_FCR_to_PDF(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'excel.path', 'excel.sheet',
                                     'excel.read_column.start_cell', 'huy.path']
        return mandatory_keys

    def automate(self):
        def process_data(sheet, full_name):
            i = 0
            with open(full_name, 'r') as file:
                for line in file:
                    i += 1
                    if i >= 4:
                        tam = line.strip()
                        sheet.cells(i - 2, 1).value = tam
                        if i == 5:
                            sheet.cells(i - 2, 1).value = tam[:80]
                        if tam.strip()[:4] in ["1`1B", "`1B", "1 `1"]:
                            sheet.cells(i - 2, 1).value = ""
                        if tam.strip()[:11] == "Attachment":
                            sheet.cells(i - 3, 1).select()
                            xw.Range(selection, selection).api.EntireRow.PageBreak = True
                            # Assuming Copy_WaterMark function does similar operation
                            # Call Copy_WaterMark(i + 25, "FCR")
                        if tam.strip()[:7] == "Receipt":
                            sheet.cells(i - 2, 1).value = tam[:34]

        # Function to handle the main logic
        def fcr_generator_sh_cnee_fcr():
            wb = xw.Book.caller()
            sheet = wb.sheets['print']
            filepath = sheet.range('H2').value
            if not os.path.exists(filepath):
                print("Folder", filepath, "does not exist.")
                return

            # Define other variables like vSH_CO, vFCR, vOCB, etc.

            f = 1
            for file_name in os.listdir(filepath):
                if file_name.lower().endswith('.txt'):
                    full_name = os.path.join(filepath, file_name)
                    fname = file_name[:-4]  # Remove the '.txt' extension
                    sheet.range('A6').value = ""  # Clear cell A6
                    process_data(sheet, full_name)
                    # Other processing steps like CopyFCR, Export_PDF, etc.

                    # Delete the text file
                    os.remove(full_name)

                f += 1

            sheet.range('A6').value = ""  # Clear cell A6

            # Show a completion message
            print("(*^_^*) Completed (*^_^*)")

        # This line allows the function to be called from Excel
        xw.Book('path_to_your_excel_file.xlsx').macros[fcr_generator_sh_cnee_fcr].module

        # Assuming NUM_On is a function to turn on something, it's not clear from the provided code
        # You can implement similar functionality as needed

    def delete_ws(fname):
        wb = xw.Book.caller()
        wb.app.display_alerts = False
        try:
            wb.sheets[fname].delete()
        except:
            pass
        wb.app.display_alerts = True

    def create_ws(fname):
        wb = xw.Book.caller()
        screen_update_state = wb.app.screen_updating
        status_bar_state = wb.app.display_status_bar
        calc_state = wb.app.calculation
        events_state = wb.app.enable_events
        display_page_break_state = wb.sheets.active.display_page_breaks

        wb.app.screen_updating = False
        wb.app.display_status_bar = False
        wb.app.calculation = 'manual'
        wb.app.enable_events = False
        wb.sheets.active.display_page_breaks = False

        try:
            sheet = wb.sheets[fname]
            sheet.cells.clear()
        except:
            wb.sheets.add().name = fname

        sheet = wb.sheets[fname]
        sheet.cells.clear()
        sheet.activate()
        sheet.range('A1:Z20000').number_format = '@'

        wb.app.screen_updating = screen_update_state
        wb.app.display_status_bar = status_bar_state
        wb.app.calculation = calc_state
        wb.app.enable_events = events_state
        wb.sheets.active.display_page_breaks = display_page_break_state

    def copy_fcr(fname, v_type_p):
        wb = xw.Book.caller()
        sheet_count = len(wb.sheets)
        flag = False
        for i in range(1, sheet_count + 1):
            if wb.sheets[i].name == v_type_p:
                flag = True
                break

        if not flag:
            print("Cannot find template " + v_type_p + ", please add Template")
            return

        wb.sheets[v_type_p].api.Copy(After=wb.sheets[sheet_count].api)
        wb.sheets.active.name = fname
        wb.sheets['print'].activate()

    def read_txt_file():
        wb = xw.Book.caller()
        my_file = "C:\\Users\\onk004\\OneDrive - Maersk Group\\MACRO\\OPEX\\006 - PDF FCR\\New macro with ESF\\FCR\\MLSGNNKO.TXT"
        with open(my_file, 'r') as file:
            text = file.read()
            pos_lat = text.find("latitude")
            pos_long = text.find("longitude")
            wb.sheets['Sheet1'].range('A1').value = text[pos_lat + 10: pos_lat + 15]
            wb.sheets['Sheet1'].range('A2').value = text[pos_long + 11: pos_long + 16]

    def process_files(file_name):
        wb = xw.Book.caller()
        wb.app.screen_updating = False
        os.system("Notepad.exe " + file_name)
        xw.apps.active.api.SendKeys("^(a)", 1)
        xw.apps.active.api.Wait(1000)
        while True:
            try:
                xw.apps.active.api.SendKeys("^(c)", 1)
                xw.apps.active.api.Wait(1000)
            except:
                break

        wb.sheets['temp'].activate()
        xw.apps.active.api.SendKeys("^v", 1)
        os.system("taskkill /f /im notepad.exe")
        wb.app.screen_updating = True
