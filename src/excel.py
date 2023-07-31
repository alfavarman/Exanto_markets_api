import logging
from typing import List, Dict, Any


# logger to log to console with information
logging.getLogger().setLevel(logging.INFO)


class ExelHandler:
    # tickets not from exanto, none for empty cell
    EXCLUDED_TICKERS = ("BTC", "ETH", "TLT", "BUSD", None)
    CASH = "Gotowka na koncie"
    COLUMN_E = 5
    COLUMN_A = 1

    def __init__(self, file_name: str, load_workbook_func):
        self.file_name = file_name
        self.load_workbook_func = load_workbook_func
        self.file = None
        self.positions = None

    def __enter__(self):
        self.file = self.load_workbook_func(self.file_name)
        self.positions = self._load_file_positions()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.file.save(self.file_name)
        self.file.close()
        logging.info(f"File {self.file_name} updated and closed successfully")

    def update_file_with_api_data(self, api_data: dict) -> None:
        """ update file with api data"""
        for position in api_data["positions"]:
            self._update_position(api_position=position)
        self._update_available_cash(api_currencies=api_data["currencies"])

    def _load_file_positions(self) -> dict:
        """
        Create a dictionary of positions from the Excel file.
        The key is the "ticker" of the position, and the value is the row index in the Excel sheet.

        :return: A dictionary representing positions with tickers as keys and row indexes as values.
        :raises: ValueError if there is a duplicated ticker in the Excel sheet.
        """
        positions = {}
        encountered_tickers = set()

        for row_number, row_data in enumerate(self.file.active.iter_rows(min_row=self.COLUMN_A, values_only=True),
                                              start=1):
            ticker = row_data[1]
            # Skip excluded tickers
            if ticker in self.EXCLUDED_TICKERS or row_data[0] == self.CASH:
                continue

            # Check for duplicated ticker (mistake in Excel)
            if ticker in encountered_tickers:
                raise ValueError(f"Duplicated ticker in Excel sheet: {ticker}")

            positions[ticker] = row_number
            encountered_tickers.add(ticker)

        return positions

    @staticmethod
    def _repr_in_thousands(value: str) -> float:
        """
        Represent the value in thousands.
        For instance, "13697.75" -> 13.7.

        :param value: The value to represent as a string.
        :return: The value represented in thousands as a float.
        :raises: ValueError if the input is not a valid number.
        """
        try:
            return round(float(value) / 1000, 1)
        except ValueError as e:
            raise ValueError(f"{value} is not a valid number. {e}")

    def _update_position(self, api_position: dict) -> None:
        """
        Update the position with values from the API.

        :param api_position: A dictionary containing position details from the API.
        """
        ticker = api_position["symbolId"].split(".")[0]
        if ticker not in self.positions:
            logging.warning(f"Ticker {ticker} not found in Excel file")
        else:
            value = self._repr_in_thousands(api_position["convertedValue"])
            self._update_row_value(ticker=ticker, value=value)

    def _update_available_cash(self, api_currencies: List[Dict[str, Any]]) -> None:
        """
        Update available cash by accumulating cash values of all currencies held by the user.
        NOTE: The value is provided in the main currency of the exanto account.

        :param api_currencies: A list of dictionaries representing currencies and their converted values.
                               Each dictionary should have "convertedValue" as one of the keys.
        """
        available_cash = sum(self._repr_in_thousands(value=currency["convertedValue"]) for currency in api_currencies)
        self._update_row_value(ticker="USD", value=available_cash)

    def _update_row_value(self, ticker: str, value: float) -> None:
        """
        Update the value for the position of a given ticker.

        :param ticker: The ticker symbol of the position.
        :param value: The value to be set for the position. If value is 0, the Excel cell will be emptied.
        """
        self.file.active.cell(row=self.positions[ticker], column=self.COLUMN_E).value = value or None
