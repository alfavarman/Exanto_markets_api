import sys

from api_services import get_api_data
from excel import ExelHandler


def main():
    file_name = sys.argv[1] if len(sys.argv) == 2 else "portfel.xlsx"

    # get API data
    api_data = get_api_data()
    file = ExelHandler(file_name=file_name)
    file.update_file_with_api_data(api_data=api_data)


if __name__ == "__main__":
    main()
