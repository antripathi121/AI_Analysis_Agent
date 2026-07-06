import pandas as pd


class ExcelReader:
    """
    Reads the Sales Report and extracts data
    from the 'All Brands Weekly' sheet.
    """

    SHEET_NAME = "All Brands Weekly"

    def __init__(self, file_path):
        self.file_path = file_path
        self.sheet = None

    def load_sheet(self):
        """
        Load only the 'All Brands Weekly' sheet.
        """
        self.sheet = pd.read_excel(
            self.file_path,
            sheet_name=self.SHEET_NAME,
            header=None
        )

        return self.sheet