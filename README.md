# Update Execl Portfolio Exanto API

## Overview
This is a simple Python script to update an Excel portfolio file with data obtained from the EXANTO API. The script is designed to work with an Excel file containing financial positions and available cash. The data from the API is used to update the values in the corresponding positions.


## Prerequisites
- Python 3.x

## Virtual Environment (venv)
To ensure a clean and isolated environment for running this project, it is recommended to use a virtual environment. Here's how to set it up:

1. Open a terminal or command prompt and navigate to the project directory:
```commandline
cd path_to/project/directory/
```

2. Create a new virtual environment by running the following command:
```commandline
python -m venv venv
```

3. Activate the virtual environment:
 On macOS and Linux:
  ```
  source venv/bin/activate
  ```
4. Once the virtual environment is activated, you can proceed with the next steps.

## Configuration
Before running the script, you need to configure the API access and set up the environment variables. Here's how to do it:

### Edit .env-TEMPLATE file
1. fileReplace the placeholder values (`XYZ0000.000`, `UUiD`, `token`, `USD`, `3.0`) with your actual API credentials and configuration provided by the API provider.
2. rename it to .env

## API Service
The `api_service.py` file contains functions to fetch data from the API. It uses the `requests` library to make HTTP requests and the `dotenv` library to load environment variables from the `.env` file.

## Excel File Format
The script assumes that the Excel file follows a specific format with the following columns:
1. Column A: Position name (string) - edit by user
2. Column B: Ticker symbol (string) - edit by user
3. Column C: Quantity in % (float or int) - autogenerate by execl
4. Column D: Profitable (colors: red, green, yellow) - edit by user
5. Column E: Deposition value in USD represented in thousands e.x. 17,4 = 17,4XX.XX (float)

**Note:** The script will skip positions with tickers listed in `EXCLUDED_TICKERS` (defined in `ExelHandler` class) and the position "Gotowka na koncie" (CASH). Should You change those in execl you should change it aslo in ExeclHandler class.


## Usage

1. Install the required Python libraries:

```
pip install openpyxl
```

2. Run the main script using the command:

```
python main.py [file_path]
```

If the `file_path` is not provided, the default file path "portfel.xlsx" will be used.

## License
### Creative Commons Attribution-NonCommercial (CC BY-NC) license.

 This project is licensed under the Creative Commons Attribution-NonCommercial (CC BY-NC) license.

You are free to:

- Share: copy and redistribute the material in any medium or format.
- Adapt: remix, transform, and build upon the material.

Under the following terms:

- Attribution: You must give appropriate credit, provide a link to the license, and indicate if changes were made. You may do so in any reasonable manner, but not in any way that suggests the licensor endorses you or your use.
- NonCommercial: You may not use the material for commercial purposes.

This is a human-readable summary of (and not a substitute for) the [full legal text of the CC BY-NC license](https://creativecommons.org/licenses/by-nc/4.0/legalcode).

