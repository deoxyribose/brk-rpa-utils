import getpass
import json
import os
import subprocess
import time
from pathlib import Path

import win32com.client  # pywin32
from dotenv import load_dotenv  # python-dotenv
from loguru import logger as log
from playwright.sync_api import Playwright

load_dotenv()
sapshcut_path = Path(os.getenv("SAPSHCUT_PATH"))
pam_path = os.getenv("PAM_PATH")
ri_path = os.getenv("RI_PATH")
robot_name = getpass.getuser()


def get_credentials(robot_name, fagsystem):
    pass_file = Path(pam_path) / robot_name / f"{robot_name}.json"

    try:
        with open(pass_file) as file:
            json_string = json.load(file)

        username = json_string[fagsystem]["username"]
        password = json_string[fagsystem]["password"]

        return username, password

    except FileNotFoundError:
        log.error("File not found", exc_info=True)
    except json.JSONDecodeError:
        log.error("Invalid JSON in file", exc_info=True)
    except Exception:
        log.error("An error occurred:", exc_info=True)

    return None, None


def start_opus(robot_name):
    """
    Starts Opus using sapshcut.exe and credentials from PAM.

    The robot00X.json file has the structure:

    {
    "ad": { "username": "robot00X", "password": "x" },
    "opus": { "username": "jrrobot00X", "password": "x" },
    "rollebaseretindgang": { "username": "jrrobot00X", "password": "x" }
    }
    """

    # unpacking
    username, password = get_credentials(robot_name, fagsystem="opus")

    if not username or not password:
        log.error("Failed to retrieve credentials for robot", exc_info=True)
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
        log.error("Failed to start SAP session", exc_info=True)
        return None


def start_ri(playwright: Playwright) -> None:
    username, password = get_credentials(robot_name, fagsystem="rollebaseretindgang")

    if not username or not password:
        log.error("Failed to retrieve credentials for robot", exc_info=True)
        return None

    try:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 2560, "height": 1440})
        page = context.new_page()
        page.goto(ri_path)
        page.get_by_placeholder("Brugernavn").click()
        page.get_by_placeholder("Brugernavn").fill(username)
        page.get_by_placeholder("Brugernavn").press("Tab")
        page.get_by_placeholder("Password").click()
        page.get_by_placeholder("Password").fill(password)
        page.get_by_role("button", name="Log på").click()
        page.get_by_text("Lønsagsbehandling").click()

        return page, context, browser

    except Exception:
        log.error("An error occurred while logging into the portal", exc_info=True)
        return None


def main():
    get_credentials(robot_name)
    start_opus(robot_name)
    start_ri()


if __name__ == "__main__":
    main()
