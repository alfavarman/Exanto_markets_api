# list of tickers to skip
EXCLUDED_TICKERS = ("BTC", "ETH", "USD", "TLT", "BUSD")
# cell description in execl
AVAILABLE_CASH = "Dolar amerykański (gotówka) Exante"
# ticker column
COL_B = 2
# exposition column
COL_E = 5


def get_positions_dict_from_sheet(sheet, ticker_column: int) -> dict:
    """
    get dictionary key value pairs
    key: str representing ticker of position
    value: number representing exel row
    """
    # last_row_with_data = sheet.max_row

    positions = {}
    for index, row in enumerate(
        sheet.iter_rows(ticker_column, values_only=True), start=1
    ):
        # skip excluded tickers
        if row[1] in EXCLUDED_TICKERS:
            continue

        try:
            positions[row[1]] = index
        except KeyError:
            print(f"Ticker: {row[1]} duplicated in exel sheet")

    return positions


def validate_file_name(name: str, file_types: tuple) -> bool:
    """validate file type"""
    file_type = name.split(".")[-1]
    return file_type in file_types


def represent_in_thousands(value: float) -> float:
    """represent value in thousands"""
    try:
        float_value = float(value)
        result = float_value / 1000
        return round(result, 1)
    except ValueError:
        print(f"ValueError: {value} is not a valid number")


def convert_to_float(value: str) -> float:
    """convert value to float None converts to 0.0"""
    if not value:
        return 0.0
    try:
        float_value = float(value)
        return float_value
    except ValueError:
        print(f"ValueError: {value} is not a valid number")


def update_file_positions_with_api_data(api_positions: dict, sheet) -> None:
    """update position to file"""

    file_positions = get_positions_dict_from_sheet(sheet, ticker_column=COL_B)
    file_tickers = file_positions.keys()

    for position in api_positions:
        ticker = position["symbolId"].split(".")[0]

        # confirm ticker exist in file
        if ticker in file_tickers:
            position_value = represent_in_thousands(position["convertedValue"])
            # confirm position is not recently sold
            cell_value = convert_to_float(
                sheet.cell(row=file_positions[ticker], column=COL_E).value
            )
            if position_value != cell_value:
                sheet.cell(row=file_positions[ticker], column=COL_E).value = (
                    position_value if position_value else None
                )
                print(f"UPDATE: {ticker} successful")
        else:
            print(f"ALERT: Ticker {ticker} not found in file")


def update_file_exante_available_cash(available_cash: float, sheet) -> None:
    """
    update available cash to file,
    function update if values are not same,
    if value is 0.0 cell will be empty
    """
    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        if row[0] == AVAILABLE_CASH:
            cell_value = convert_to_float(
                sheet.cell(row=row_number, column=COL_E).value
            )
            if available_cash != cell_value:
                sheet.cell(row=row_number, column=COL_E).value = (
                    available_cash if available_cash else None
                )
                print(f"UPDATE: {AVAILABLE_CASH} successful")
            break
    else:
        raise Exception(f"{AVAILABLE_CASH} not found in file")


def get_currencies_converted_value(data: dict[dict]) -> float:  # TODO comprehentions
    totals = 0.0
    for currency in data:
        totals += float(currency["convertedValue"])
    return represent_in_thousands(totals)
