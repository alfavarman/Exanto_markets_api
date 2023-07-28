import sys

from openpyxl import load_workbook

from src.api_services import get_api_data
from src.excel import ExelHandler


def main():
    file_name = sys.argv[1] if len(sys.argv) == 2 else "portfel.xlsx"

    # get API data
    api_data = get_api_data()

    file = ExelHandler(file_name=file_name)
    file.load_file_data()
    file.update_file_with_api_data(api_data=api_data)
    file.save_and_close()


if __name__ == "__main__":
    main()
