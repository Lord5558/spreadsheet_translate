"""
    Translates the spreadsheet - easy and quick
     by Mykyta Aleksandrov 2022

    How to use? That's how:

    >> ./spreadsheet_translate.py -n file.xlsx                Does everything but via the terminal


    from spreadsheet_translate import Spreadsheet             import the directory

    spreadsheet = Spreadsheet("file.xlsx")                    Open the workbook
    spreadsheet.translate()                                   Translate all the spreadsheets
    spreadsheet.save()                                        Save the workbook in the new file
"""

import logging
import argparse
from openpyxl import load_workbook
import transl


LEVEL = logging.INFO
FMT = '[%(levelname)s] %(asctime)s - %(message)s'
logging.basicConfig(level=LEVEL, format=FMT)

parser = argparse.ArgumentParser()
parser.add_argument("-n", "--name", type=str, required=True, help="The name of the Excel file")
args = parser.parse_args()


class Spreadsheet:

    """Opens, translates and saves the workbook (.xlsx)"""

    def __init__(self, filename_: str):
        self.workbook = None
        self.filename = filename_

    @staticmethod
    def translate_cell(text: str) -> float | str:

        """Translates individual cells after selecting non-numerical cells"""

        try:
            float(text)
        except ValueError:
            if text[0] != "=":
                return transl.translate(text)
            return text
        except TypeError:
            return ''
        else:
            return round(float(text), 2)

    def transl_sheet(self, sheet: any) -> None:

        """Iterates through each cell in each row in each column and extracts its value"""

        for col_n, column in enumerate(sheet.iter_cols(values_only=True), start=1):
            for row_n, cell in enumerate(column, start=1):
                try:
                    sheet.cell(row=row_n, column=col_n).value = self.translate_cell(cell)
                except AttributeError:
                    pass
        logging.info('Worksheet "%s" - Translation complete. Moving on...', sheet.title)
        return sheet

    def translate(self) -> list[None] | None:

        """Loads the workbook and iterates through each spreadsheet inside it before translation"""

        logging.info('Loading the workbook.')
        try:
            self.workbook = load_workbook(filename=self.filename)
        except NameError:
            logging.error("Error occurred. Couldn't load the workbook. "
                          "Possibly wrong filename.")
            return None
        else:
            logging.info('Successfully loaded the workbook.')
            return [self.transl_sheet(sheet) for sheet in self.workbook]    # The iteration is here

    def save(self):

        """Saves the workbook into a new .xlsx file"""

        name = f'Translated - {self.filename}'
        logging.info('Saving the workbook into "%s"', name)
        self.workbook.save(filename=name)
        logging.info('Done.')


if __name__ == "__main__":
    spreadsheet = Spreadsheet(args.name)
    spreadsheet.translate()
    spreadsheet.save()
