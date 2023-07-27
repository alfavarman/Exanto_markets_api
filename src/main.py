import sys

from openpyxl import load_workbook

from api_services import get_api_data
from utils import (
    validate_file_name,
    update_file_positions_with_api_data,
    update_file_exante_available_cash, get_currencies_converted_value
)

# list of accepted files types
FILES_TYPES = ("xlsx", "xls")


def main():
    # ensure proper usage: get command line arg or exit
    if len(sys.argv) != 2:
        sys.exit("Usage: python extante.py portfolio.xlsx \n"
                 "Info: README.md")

    file_name = sys.argv[1]

    if not validate_file_name(file_name, FILES_TYPES):
        sys.exit(f"argument {file_name} is not a valid file type"
                 f"valid files type: {FILES_TYPES}"
                 "Info: README.md")

    # get API data
    api_data = get_api_data()

    # open file
    try:
        file = load_workbook(f'{file_name}')
    except Exception as e:
        # FileNotFoundError or LoadError
        sys.exit(f"Error: {e}")
    else:
        sheet = file.active

    # update file
    api_positions = api_data["positions"]

    available_cash = get_currencies_converted_value(data=api_data["currencies"])
    update_file_positions_with_api_data(api_positions=api_positions, sheet=sheet)
    update_file_exante_available_cash(available_cash=available_cash, sheet=sheet)

    # save the file
    file.save(file_name)
    print(f"FILE >{file_name}< update successful")


if __name__ == "__main__":
    main()
