import subprocess
import sys
import time

from pywinauto import Application


EXE_PATH = r"C:\Users\K1022108\Downloads\LaunchScript\DesignAutomationScriptLauncher.exe"
WINDOW_TITLE = "Design Automation Hub"
TOOLBAR_TO_LAUNCH = "DDP Toolbar"


def open_launcher():
    subprocess.Popen([EXE_PATH])


def wait_for_window(timeout_sec=30):
    deadline = time.time() + timeout_sec
    last_error = None

    while time.time() < deadline:
        try:
            app = Application(backend="uia").connect(title=WINDOW_TITLE, timeout=2)
            window = app.window(title=WINDOW_TITLE)
            window.wait("visible enabled ready", timeout=5)
            return window
        except Exception as ex:  # noqa: BLE001
            last_error = ex
            time.sleep(0.5)

    raise RuntimeError(f"Could not find '{WINDOW_TITLE}'. Last error: {last_error}")


def select_toolbar(window, toolbar_name):
    radio = window.child_window(title=toolbar_name, control_type="RadioButton").wrapper_object()
    radio.click_input()


def click_launch(window):
    launch_button = window.child_window(title="Launch", control_type="Button").wrapper_object()
    launch_button.click_input()


def main():
    try:
        open_launcher()
        window = wait_for_window()
        select_toolbar(window, TOOLBAR_TO_LAUNCH)
        click_launch(window)
        print(f"Launched: {TOOLBAR_TO_LAUNCH}")
        return 0
    except Exception as ex:  # noqa: BLE001
        print(f"ERROR: {ex}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
