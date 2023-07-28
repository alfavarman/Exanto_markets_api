import logging
from typing import List, Dict, Any
from openpyxl import load_workbook


class ExelHandler:
    EXCLUDED_TICKERS = ("BTC", "ETH", "TLT", "BUSD")
    CASH = "Gotowka na koncie"

    def __init__(self, file_name: str):
        self.file_name = file_name
        self.sheet = load_workbook(self.file_name).active
        self.positions = self._load_file_positions()

    def update_file_with_api_data(self, api_data: dict) -> None:
        for position in api_data["positions"]:
            self._update_position(api_position=position)
        self._update_available_cash(api_currencies=api_data["currencies"])
        self.sheet.save(self.file_name)
        self.sheet.close()
        logging.info(f"File {self.file_name} updated successfully")

    def _load_file_positions(self):
        positions = {}
        for row_number, row_data in enumerate(self.sheet.iter_rows(2, values_only=True), start=1):
            # skip excluded tickers
            if row_data[1] in self.EXCLUDED_TICKERS or row_data[0] == self.CASH:
                continue
            if row_data[1] in positions:
                raise Exception(f"Ticker: {row_data[1]} duplicated in exel sheet")
            positions[row_data[1]] = row_number
        return positions

    def _repr_in_thousands(self, value: str) -> float:
        """represent value in thousands"""
        try:
            return round((float(value) / 1000), 1)
        except ValueError:
            print(f"ValueError: {value} is not a valid number")

    def _update_position(self, api_position: dict) -> None:
        """update position to file"""
        ticker = api_position["symbolId"].split(".")[0]
        if ticker not in self.positions:
            raise Exception(f"{ticker} not found in file")
        value = self._repr_in_thousands(value=api_position["convertedValue"])
        self._update_row_value(ticker=ticker, value=value)

    def _update_available_cash(self, api_currencies: List[Dict[str, Any]]) -> None:
        available_cash = sum(self._repr_in_thousands(value=currency["convertedValue"]) for currency in api_currencies)
        self._update_row_value(ticker="USD", value=available_cash)

    def _update_row_value(self, ticker: str, value: float) -> None:
        self.sheet.cell(row=self.positions[ticker], column=5).value = value or None
