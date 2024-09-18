"""
File: app-manager.py
Author: Chuncheng Zhang
Date: 2024-08-26
Copyright & Email: chuncheng.zhang@ia.ac.cn

Purpose:
    Amazing things

Functions:
    1. Requirements and constants
    2. Function and class
    3. Play ground
    4. Pending
    5. Pending
"""


# %% ---- 2024-08-26 ------------------------
# Requirements and constants
import sys
import time
import pyvda
import argparse
import win32gui

from loguru import logger
from rich import print, inspect


# %% ---- 2024-08-26 ------------------------
# Function and class

def get_all_desktops():
    '''
    https://github.com/Ciantic/VirtualDesktopAccessor
    '''
    desktops = pyvda.get_virtual_desktops()
    return desktops


def get_title(app: pyvda.AppView) -> str:
    '''
    Get the title of the application.

    Args:
        - app: The application to switch to.

    Returns:
        - (str): The title of the application.
    '''
    return win32gui.GetWindowText(app.hwnd)


def switch_to_app(app: pyvda.AppView, dry_run: bool = False):
    '''
    Switch to the application.

    Args:
        - app: The application to switch to.
        - dry_run (bool, default by False): Whether to prevent switching to the application.
    '''
    title = get_title(app)
    name = app.desktop.name

    if dry_run:
        print(
            f'Dry run switch to desktop: "{name}", title: "{title}"')
        return

    # Switch to the desktop
    app.desktop.go()
    # Focus to the app
    app.set_focus()
    print(f'Switched to desktop: "{name}", title: "{title}"')
    return


# %% ---- 2024-08-26 ------------------------
# Play ground
if __name__ == "__main__":
    # desktops = get_all_desktops()
    # print(f'desktops: {desktops}')

    # ----------------------------------------
    # ---- Parse the arguments ----
    parser = argparse.ArgumentParser(
        'App management', description='Manage application automatically in windows 11                                     ')
    parser.add_argument(
        '-d', '--dry-run', help='Dry run for the testing functional.', action='store_true')
    parser.add_argument(
        '-e', '--execute', help='Actually execute the script.', action='store_true')
    parser.add_argument(
        '-t', '--title', help='Set the title to be switch to.')
    opt = parser.parse_args()

    print(f'Parsed arguments: {opt}')

    # Exit the script if not execute
    if not (opt.execute or opt.dry_run):
        logger.warning(
            'Not executing the script since the --execute argument is not set.')
        sys.exit(1)

    # Get the arguments
    dry_run = opt.dry_run
    selected_title = opt.title

    if selected_title:
        dry_run = False
        logger.debug(
            f'Specifying title: {selected_title}, the dry_run argument is set to False ignoring the argument.')

    # ----------------------------------------
    # ---- Execute ----
    # Get all the applications
    applications = pyvda.get_apps_by_z_order(current_desktop=False)

    # If opt.title exists,
    # filter the applications according to its title
    if selected_title:
        applications = [
            app for app in applications if get_title(app) == selected_title]

    # Remember the current application
    current_application = pyvda.AppView.current()
    # Set the interval
    interval = 0.2 if not dry_run else 0.0  # Seconds
    # Walk through all the applications
    for app in applications:
        try:
            switch_to_app(app, dry_run=dry_run)
            time.sleep(interval)
        except Exception:
            logger.error(f'Failed to find window: {app}, {app.hwnd}')

    # Comeback
    if not selected_title:
        switch_to_app(current_application, dry_run=dry_run)
        logger.debug(f'Restoring to the application: {current_application}')

    sys.exit()

# %% ---- 2024-08-26 ------------------------
# Pending


# %% ---- 2024-08-26 ------------------------
# Pending
