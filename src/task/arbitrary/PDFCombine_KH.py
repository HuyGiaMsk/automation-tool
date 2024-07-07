import os
from logging import Logger
from typing import Callable

from PyPDF2 import PdfMerger

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.StringUtil import get_row_index_from_excel_cell_format
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


# noinspection PyPackageRequirements
class PDFCombine_KH(AutomatedTask):
    tax_to_bill: dict[str, str] = {}

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'folder_payment_slip.folder', 'folder_wy.folder',
                                     'folder_inv.folder', 'folder_cheque_request.folder', 'excel.column.bill',
                                     'folder_combine.folder', 'excel.column.wht']
        return mandatory_keys

    def automate(self):
        logger: Logger = get_current_logger()

        excel_reader: ExcelReaderProvider = XlwingProvider()

        path_to_excel_contain_pdfs_content = self._settings['excel.path']
        workbook = excel_reader.get_workbook(path=path_to_excel_contain_pdfs_content)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        worksheet = excel_reader.get_worksheet(workbook, sheet_name)

        bills: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                 self._settings['excel.sheet'],
                                                                 self._settings[
                                                                     'excel.column.bill'])

        withholding_taxes: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                             self._settings['excel.sheet'],
                                                                             self._settings['excel.column.wht'])
        if len(bills) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        if len(bills) != len(withholding_taxes):
            raise Exception("Please check your input data length of bills, type_bill are not equal")

        """Step 1: Store bill-to-info mapping"""
        index: int = 0
        for tax in withholding_taxes:
            self.tax_to_bill[tax] = (bills[index])
            index += 1

        excel_row_index: int = get_row_index_from_excel_cell_format(self._settings['excel.column.bill'])

        # Loop through files in the WY folder, rename them according to the bill numbers
        self.rename_files_in_folder_wy()

        """"Step 2: Process each bill + progress bar"""

        self.current_element_count = 0
        self.total_element_size = len(bills)
        for bill in bills:
            logger.info("Processing: " + bill)
            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            self.combine_pdfs_into_one(bill, worksheet, excel_row_index)
            excel_row_index += 1

            self.current_element_count = self.current_element_count + 1
            excel_reader.save(workbook=workbook)

        self.combine_pdfs_combined()

        # Close the Excel workbook
        excel_reader.close(workbook=workbook)
        excel_reader.quit_session()
        logger.info('Done input to excel file')

    def rename_files_in_folder_wy(self):
        folder_wy: str = self._settings['folder_wy.folder']
        for file_name in os.listdir(folder_wy):
            if file_name.endswith(".pdf"):
                wy_number = file_name.split(".pdf")[0]

                bill_number = self.tax_to_bill.get(wy_number)
                if bill_number is not None:
                    os.rename(os.path.join(folder_wy, file_name),
                              os.path.join(folder_wy, f"{bill_number}_{wy_number}.pdf"))

    def combine_pdfs_into_one(self, bill, worksheet, excel_row_index):
        """
        Combine PDFs related to the given bill into one file and update the Excel sheet.
        """

        logger: Logger = get_current_logger()

        # Define folders
        folder_cheque_request: str = self._settings['folder_cheque_request.folder']
        folder_payslip: str = self._settings['folder_payment_slip.folder']
        folder_wy: str = self._settings['folder_wy.folder']
        folder_inv: str = self._settings['folder_inv.folder']

        # Step 3: Combine PDFs
        merger = PdfMerger()
        for folder in [folder_cheque_request, folder_payslip, folder_wy, folder_inv]:
            pdf_files = [file for file in os.listdir(folder) if file.startswith(bill)]
            for pdf_file in pdf_files:
                merger.append(os.path.join(folder, pdf_file))

        # Save combined PDF
        output_folder: str = self._settings['folder_combine.folder']
        output_file = os.path.join(output_folder, f"{bill}.pdf")

        logger.info("Combined {}.pdf".format(bill))
        if os.path.exists(output_file):
            os.remove(output_file)

        with open(output_file, 'wb') as output:
            merger.write(output)

        # Update Excel sheet
        counts = []
        for folder in [folder_payslip, folder_wy, folder_inv, folder_cheque_request]:
            count = self.find_and_count_pdfs(folder, bill)
            counts.append(count)
        self.update_excel_sheet(worksheet=worksheet, row_index=excel_row_index, counts=counts)

    def combine_pdfs_combined(self):
        """
        Combine PDFs combined before in each folder into one file
        """
        logger: Logger = get_current_logger()

        folder_to_combine: str = self._settings['folder_combine.folder']
        merger = PdfMerger()

        # Iterate through the PDF files in the folder
        for file in os.listdir(folder_to_combine):
            if file.lower().endswith('.pdf'):
                pdf_file_path = os.path.join(folder_to_combine, file)
                merger.append(pdf_file_path)

        # Save combined PDF
        output_folder: str = self._settings['folder_combine.folder']
        output_file = os.path.join(output_folder, "Combined.pdf")

        # Remove existing output file if exists
        if os.path.exists(output_file):
            os.remove(output_file)

        with open(output_file, 'wb') as output:
            merger.write(output)

        merger.close()

        # Delete all PDF files except the final combined one
        for file in os.listdir(folder_to_combine):
            if file.lower().endswith('.pdf') and file != "Combined.pdf":
                file_path = os.path.join(folder_to_combine, file)
                try:
                    os.remove(file_path)
                except Exception as e:
                    logger.error(f"Error deleting file {file_path}: {e}")

        logger.info('Removed old files')

    def find_and_count_pdfs(self, folder, prefix):
        """
        Find and count PDFs in the given folder with the specified prefix.
        """
        pdf_files = [file for file in os.listdir(folder) if file.startswith(prefix)
                     and (file.endswith('.pdf') or file.endswith('.PDF'))]
        return len(pdf_files)

    def update_excel_sheet(self, worksheet, row_index, counts):
        """
        Update the Excel sheet with the counts of PDFs.
        Args:
            worksheet (xlwings main.Sheet): The Excel worksheet to update.
            row_index (int): The row index where you want to update the counts.
            counts (list): A list containing the counts of PDFs for each folder.
        """
        # Define the starting column index for inputting counts
        column_index = 3  # Start from column C

        for count in counts:
            worksheet.range(row_index, column_index).value = count
            column_index += 1
