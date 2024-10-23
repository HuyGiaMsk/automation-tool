import os
from logging import Logger
from typing import Callable

import pdfplumber
from pdfplumber import PDF

from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


class PDFRead(AutomatedTask):
    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'folder_docs.folder']
        return mandatory_keys

    def clear_worksheet(self, worksheet):
        excel_reader: ExcelReaderProvider = XlwingProvider()
        used_range = worksheet.used_range
        used_range.clear_contents()

    def automate(self):
        logger: Logger = get_current_logger()

        excel_reader: ExcelReaderProvider = XlwingProvider()

        path_to_excel_contain_pdfs_content = self._settings['excel.path']
        workbook = excel_reader.get_workbook(path=path_to_excel_contain_pdfs_content)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        worksheet = excel_reader.get_worksheet(workbook, sheet_name)
        self.clear_worksheet(worksheet)

        path_to_docs = self._settings['folder_docs.folder']
        pdf_counter: int = 1

        for root, dirs, files in os.walk(path_to_docs):
            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            # files = []

            for current_pdf in files:
                if not current_pdf.lower().endswith(".pdf"):
                    continue

                current_pdf_path = os.path.join(root, current_pdf)

                try:
                    pdf: PDF = pdfplumber.open(current_pdf_path)
                except Exception as e:
                    logger.error(f"Failed to open {current_pdf_path}: {e}")
                    continue  # Skip this PDF if there's an error

                logger.info("File name : {} PDF counter  = {}".format(current_pdf, pdf_counter))
                excel_reader.change_value_at(worksheet=worksheet, row=1, column=pdf_counter, value=current_pdf)

                current_page_in_current_pdf = 2
                for number, pageText in enumerate(pdf.pages):
                    raw_text = pageText.extract_text()
                    clean_text = raw_text.replace("\x00", "").replace("=", "")

                    # print( "text at page {} : {}".format(current_page_in_current_pdf, clean_text))
                    for line in clean_text.splitlines():
                        excel_reader.change_value_at(worksheet=worksheet, row=current_page_in_current_pdf,
                                                     column=pdf_counter, value=line)
                        current_page_in_current_pdf += 1

                excel_reader.save(workbook=workbook)

                pdf_counter += 1

        excel_reader.close(workbook=workbook)
        excel_reader.quit_session()
        logger.info('Closed excel file - Done')
