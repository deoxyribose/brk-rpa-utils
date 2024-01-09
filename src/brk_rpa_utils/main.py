"""
import os
import json
import subprocess
import time
from pathlib import Path
from bs4 import BeautifulSoup  # BeautifulSoup4
import pandas as pd
import re
import io
from loguru import logger
import win32com.client  # pywin32
"""


def _get_credentials(pam_path, robot_name, fagsystem) -> None:
    """
    Internal function to retrieve credentials.

    pam_path = os.getenv("PAM_PATH")

    Define pam_path in an .env file in the root of your project. Add paths like so:
    SAPSHCUT_PATH=C:/Program Files (x86)/SAP/FrontEnd/SAPgui/sapshcut.exe

    robot_name = getpass.getuser()

    Under the pam_path uri der should be a robot_name.json file with the structure:

    {
    "ad": { "username": "robot00X", "password": "x" },
    "opus": { "username": "jrrobot00X", "password": "x" },
    "rollebaseretindgang": { "username": "jrrobot00X", "password": "x" }
    }
    """
    pass_file = Path(pam_path) / robot_name / f"{robot_name}.json"

    try:
        with open(pass_file) as file:
            json_string = json.load(file)

        username = json_string[fagsystem]["username"]
        password = json_string[fagsystem]["password"]

        return username, password

    except FileNotFoundError:
        logger.error("File not found", exc_info=True)
    except json.JSONDecodeError:
        logger.error("Invalid JSON in file", exc_info=True)
    except Exception:
        logger.error("An error occurred:", exc_info=True)

    return None, None


def start_opus(pam_path, robot_name, sapshcut_path) -> None:
    """
    Starts Opus using sapshcut.exe and credentials from PAM.

    load_dotenv()
    sapshcut_path = Path(os.getenv("SAPSHCUT_PATH"))

    The robot_name.json file should have the structure:

    {
    "ad": { "username": "robot_name", "password": "x" },
    "opus": { "username": "robot_name", "password": "x" },
    "rollebaseretindgang": { "username": "robot_name", "password": "x" }
    }
    """

    # unpacking
    username, password = _get_credentials(pam_path, robot_name, fagsystem="opus")

    if not username or not password:
        logger.error("Failed to retrieve credentials for robot", exc_info=True)
        return None

    command_args = [
        str(sapshcut_path),
        "-system=P02",
        "-client=400",
        f"-user={username}",
        f"-pw={password}",
    ]

    subprocess.run(command_args, check=False)  # noqa: S603
    time.sleep(3)

    try:
        sap = win32com.client.GetObject("SAPGUI")
        app = sap.GetScriptingEngine
        connection = app.Connections(0)
        session = connection.sessions(0)
        return session

    except Exception:
        logger.error("Failed to start SAP session", exc_info=True)
        return None


def start_ri(pam_path, robot_name, ri_url, Playwright) -> None:
    """
    Starts s browser pointed to ri_url (fx https://portal.kmd.dk/irj/portal)
    and logs into Rollebaseret Indgang (RI) using  and credentials from PAM.

    load_dotenv()
    ri_url = os.getenv("RI_PATH")

    from playwright.sync_api import Playwright

    The robot_name.json file should have the structure:

    {
    "ad": { "username": "robot_name", "password": "x" },
    "opus": { "username": "robot_name", "password": "x" },
    "rollebaseretindgang": { "username": "robot_name", "password": "x" }
    }
    """

    username, password = _get_credentials(pam_path, robot_name, fagsystem="rollebaseretindgang")

    if not username or not password:
        logger.error("Failed to retrieve credentials for robot", exc_info=True)
        return None

    try:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 2560, "height": 1440})
        page = context.new_page()
        page.goto(ri_url)
        page.get_by_placeholder("Brugernavn").click()
        page.get_by_placeholder("Brugernavn").fill(username)
        page.get_by_placeholder("Brugernavn").press("Tab")
        page.get_by_placeholder("Password").click()
        page.get_by_placeholder("Password").fill(password)
        page.get_by_role("button", name="Log på").click()
        page.get_by_text("Lønsagsbehandling").click()

        return page, context, browser

    except Exception:
        logger.error("An error occurred while logging into the portal", exc_info=True)
        return None


def parse_ri_html_report_to_dataframe(mhtml_path) -> None:
    """
    Parses an mhtml file downloaded from Rollobaseret Indgang.
    The default download calls the file xls, but it is a kind of html file.

    ## Usage
    mhtml_path = Path(folder_data_session / 'test.html')

    df_mhtml = parse_ri_html_report_to_dataframe(mhtml_path)
    """
    try:
        # Read MHTML file
        with open(mhtml_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # Find the HTML part of the file
        match = re.search(r'<html.*<\/html>', content, re.DOTALL)
        if not match:
            raise ValueError("No HTML content found in the file")
        html_content = match.group(0)

        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(html_content, 'lxml')

        # Find all tables within the parsed HTML
        tables = soup.find_all('table')
        if not tables:
            raise ValueError("No tables found in the HTML content")

        # Select the largest table based on character count
        largest_table = max(tables, key=lambda table: len(str(table)))

        # Convert the largest HTML table to a pandas DataFrame
        df_mhtml = pd.read_html(io.StringIO(str(largest_table)), decimal=',', thousands='.', header=None)
        if not df_mhtml:
            raise ValueError("Failed to parse the largest table into a DataFrame")

        df_mhtml = df_mhtml[0]
        df_mhtml.columns = df_mhtml.iloc[0]
        df_mhtml = df_mhtml.drop(0)
        df_mhtml.reset_index(drop=True, inplace=True)
        df_mhtml.rename(columns={'Slut F-periode': 'date', 'Lønart': 'lonart', 'Antal': 'antal'}, inplace=True)

        # Convert 'date' column to datetime
        df_mhtml['date'] = pd.to_datetime(df_mhtml['date'], format='%d%m%Y')
        return df_mhtml

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        return None
