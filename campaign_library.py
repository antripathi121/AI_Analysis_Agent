import pandas as pd
import os


class CampaignLibrary:
    """
    Campaign Library

    Reads and stores the Daily Redemption Report.
    """

    LIBRARY_FILE = "campaign_library.xlsx"

    REQUIRED_COLUMNS = [
        "Offer Name",
        "CPG",
        "Type",
        "Start Date",
        "End Date",
        "Unique",
        "New",
        "Total"
    ]

    def __init__(self):
        self.data = None

    def load_redemption_report(self, uploaded_file):
        """
        Read Daily Redemption Report.

        NOTE:
        Actual headers are on row 2,
        so header=1.
        """

        df = pd.read_excel(
            uploaded_file,
            header=1
        )

        df.columns = df.columns.str.strip()

        self.data = df[self.REQUIRED_COLUMNS].copy()

        return self.data

    def save_library(self):

        if self.data is None:
            raise Exception("No campaign data loaded.")

        self.data.to_excel(
            self.LIBRARY_FILE,
            index=False
        )

    def load_library(self):

        if not os.path.exists(self.LIBRARY_FILE):
            raise FileNotFoundError("Campaign library not found.")

        self.data = pd.read_excel(self.LIBRARY_FILE)

        return self.data