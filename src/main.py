import os
import sys

from api_services import get_api_data
from excel import ExelHandler
from openpyxl import load_workbook


def main():
    # if file_path name is not given use default file_path
    file_path = sys.argv[1] if len(sys.argv) == 2 else "portfel.xlsx"
    # check if file exists
    if not os.path.isfile(file_path):
        sys.exit(f"File {file_path} not found")

    # get API data
    api_data = get_api_data()

    # with context manager to handle errors
    with ExelHandler(file_name=file_path, load_workbook_func=load_workbook) as file_handler:
        file_handler.update_file_with_api_data(api_data=api_data)


if __name__ == "__main__":
    main()
