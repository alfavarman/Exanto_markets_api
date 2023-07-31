import os
import json
import requests
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth

load_dotenv(dotenv_path="../.env")


def get_env_variable(name: str):
    try:
        return os.getenv(name)
    except KeyError:
        message = f"Expected environment variable '{name}' not set."
        raise Exception(message)


id_acc_id = get_env_variable("ID_ACC_ID")
currency = get_env_variable("CURRENCY")
version = get_env_variable("VERSION")
username = get_env_variable("USERNAME")
password = get_env_variable("PASSWORD")
api_url = f'{get_env_variable("API_URL")}'


def get_api_data() -> dict:
    data = requests.get(url=api_url, auth=HTTPBasicAuth(username=username, password=password))
    return data.json()


# temp monkey-patching of get_api_data for development (no valid credential)
def get_api_data_mock() -> dict:
    with open("src/data.json", "r") as file:
        data = json.load(file)
    return data
