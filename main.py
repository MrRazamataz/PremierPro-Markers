import traceback
import urllib.request
import webbrowser

import PySimpleGUI as sg
import csv
import os
import json
import ctypes
import platform
import time

version = 6
version_string = "v0.0.6"
debug = False
template_filename = "template.json"
template = {}
markers = []
blink = {}

if not os.path.isfile("settings.json"):
    with open("settings.json", "w") as f:
        f.write(json.dumps({"dpi": True}, indent=4))
else:
    with open("settings.json", "r") as f:
        settings = json.loads(f.read())
if not os.path.isfile(template_filename):
    with open(template_filename, "w") as f:
        f.write(json.dumps({"before": "", "after": ""}, indent=4))
    sg.Popup(
        "Welcome!\nPlease create a template in the Templates tab if you want text to come before or after the chapters.",
        title="First Run")
    template = {"before": "", "after": ""}
else:
    with open(template_filename, "r") as f:
        template = json.loads(f.read())

# find other filesnames that contain "template" and ".json"
template_files = [f for f in os.listdir() if "template" in f and ".json" in f and f != template_filename]
if len(template_files) > 0:
    template_files.sort()
else:
    template_files = ["Last used"]


def check_for_update(window: sg.Window):
    global debug
    print("Checking for updates...")
    try:
        with urllib.request.urlopen(
                "https://raw.githubusercontent.com/MrRazamataz/PremierPro-Markers/master/version.json") as url:
            data = json.loads(url.read().decode())
            if data["version"] > version:
                print("Update available!")
                release_notes = data["release_notes"]
                window["-program_log-"].update(f"{version_string} (outdated - update available)")
                return True, release_notes
            elif data["version"] == version:
                print("Up to date!")
                window["-program_log-"].update(version_string)
            elif data["version"] < version:
                print("You're using a newer version than the latest release, DEBUG MODE ON!")
                print("Make sure to PR any cool changes you make :)\n----")
                debug = True
                window["-program_log-"].update(f"{version_string} (modified)")
    except Exception as e:
        print("Error checking for updates: " + str(e))
        print(traceback.format_exc())
    return False, "none"


def make_dpi_aware():
    if int(platform.release()) >= 8:
        if settings["dpi"]:
            ctypes.windll.shcore.SetProcessDpiAwareness(True)


def format_time(input_time):
    hours, minutes, seconds, frames = map(int, input_time.split(':'))
    formatted_time = "{:02d}:{:02d}".format(hours * 60 + minutes, seconds)
    return formatted_time


def open_csv_file(file_path):
    print("Opening file: " + file_path)
    line_num = 1
    try:
        with open(file_path, encoding='utf-16') as csvfile:
            reader = csv.DictReader(csvfile, delimiter='\t')
            print("File opened successfully")
            markers.clear()
            for i in reader:
                csv_time = format_time(i['In'])
                if line_num == 1:
                    if "00:00" not in csv_time: # yt needs it to start with 00:00
                        csv_time = csv_time[:4] + "00:00"
                        csv_time = csv_time.replace("00:000:00", "00:00")
                        add_to_blinker("-modified_warning-")
                    else:
                        remove_from_blinker("-modified_warning-")
                markers.append(f"{csv_time} {i['Marker Name']}")
                line_num += 1
            return True
    except Exception as e:
        print("Error opening file: " + str(e))
        print(traceback.format_exc())
        return False


def update_output(window: sg.Window):
    window["-modified_warning-"].update(visible=False)
    window["-operation_status-"].update("Updating output...")
    output = template["before"] + "\n"
    for i in markers:
        output += i + "\n"
    output += template["after"]
    window['output'].update(output)
    window["-operation_status-"].update("Output updated.")


def copy_to_clipboard(window, text):
    window.TKroot.clipboard_clear()
    window.TKroot.clipboard_append(text)


def get_setting_boolean(setting):
    return settings[setting]


def blinker(window):
    while True:
        time.sleep(0.2)
        elements = list(blink.keys())
        for element in elements:
            if blink[element]:
                visible = window[element].visible
                if not visible:
                    window[element].update(visible=True)
                current_text_color = window[element].TextColor
                current_background_color = window[element].BackgroundColor
                window[element].update(text_color=current_background_color, background_color=current_text_color)
                time.sleep(0.2)
                window[element].update(text_color=current_text_color, background_color=current_background_color)


def add_to_blinker(element):
    blink[element] = True


def remove_from_blinker(element):
    blink[element] = False


def main():
    global template
    make_dpi_aware()
    nav_menu_def = [['File', ["Open... [Ctrl + O]", "Save template [Ctrl + S]", "Exit [Ctrl + W]"]],
                    ['Edit', ["Copy output [Ctrl + C]", "Settings"]],
                    ['Help', ["GitHub link [Ctrl + U]", "Check for updates"]]]

    markers_layout = [
        [sg.Text(
            'Automagically create YouTube chapters from PremierPro markers. Made by MrRazamataz - inspired by RavinMaddHatter.',
            expand_x=True, key="-info_text-")],
        [sg.FileBrowse("Browse", key="-file_browse-", enable_events=True,
                       file_types=(("CSV File", "*.csv"),))],
        [sg.Text("Output:")],
        [sg.Multiline(key='output', expand_x=True, expand_y=True)],
        [sg.Button("Copy", key="-copy-", tooltip="Copy to clipboard [Ctrl + C]."), sg.Text("First timestamp modified to be 00:00!", text_color="red", key="-modified_warning-", background_color="black", visible=False)],
    ]
    template_layout = [
        [sg.Text('Create the template for your description generator.', expand_x=True), sg.Push(),
         sg.DropDown(template_files, key="-template_files-", enable_events=True, default_value="Last used",
                     readonly=True, auto_size_text=True, size=(20, 1))],
        [sg.Text("Before chapters:")],
        [sg.Multiline(key='-before-', expand_x=True, expand_y=True, default_text=template["before"], enable_events=True,
                      size=(20, 10))],
        [sg.Text("After chapters:")],
        [sg.Multiline(key='-after-', expand_x=True, expand_y=True, default_text=template["after"], enable_events=True,
                      size=(20, 10))],
        [sg.Text(
            "The defaut template is the last used template. If you want to save a template, enter a name in the box below and click save, this will save so you have templates you can switch between.",
            size=(90, 3))],
        [sg.Button("Save as template", key="-save-", tooltip="Save template.")],
    ]

    logs_layout = [
        [sg.Output(size=(70, 20), key='output', expand_x=True, expand_y=True)],
    ]
    layout = [
        [sg.Menu(nav_menu_def, key="-menu-")],
        [sg.TabGroup(
            [[sg.Tab('Markers', markers_layout), sg.Tab('Template', template_layout), sg.Tab('Log', logs_layout)]],
            expand_x=True, expand_y=True, enable_events=True, key="-tab_group-")],

        [sg.StatusBar("To get started, open a file", key="-operation_status-", expand_x=True, auto_size_text=False,
                      right_click_menu=["&Right", ["Open...", "Exit"]], tooltip="The status of the current operation."),
         sg.StatusBar(f"{version_string} (checking for updates...)", key="-program_log-", justification="right", auto_size_text=False,
                      tooltip="Program log.")],
    ]

    window = sg.Window('PremierPro Markers', layout, resizable=True, finalize=True)
    window.bind("<Control-o>", "Open...")
    window.bind("<Control-u>", "GitHub link")
    window.bind("<Control-s>", "-save-")
    window.bind("<Control-w>", "Exit")
    window.bind("<Control-c>", "-copy-")
    window.perform_long_operation(lambda: blinker(window), '-blinker_thread-')
    window.perform_long_operation(lambda: check_for_update(window), "-thread-")
    while True:
        event, values = window.read()
        if debug:
            print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            with open(template_filename, "w") as f:
                f.write(json.dumps(template, indent=4))
            break
        elif "Open..." in event:
            window["-operation_status-"].update("Select a file...")
            file = sg.popup_get_file('Open file', no_window=True, file_types=(("CSV File", "*.csv"),))
            if file:
                if open_csv_file(file) == False:
                    sg.popup_error("Error opening file")
                update_output(window)
        elif event == '-file_browse-':
            window["-operation_status-"].update("Select a file...")
            if values['-file_browse-']:
                if open_csv_file(values['-file_browse-']) == False:
                    sg.popup_error("Error opening file")
                update_output(window)
        elif event == "-before-":
            template["before"] = values["-before-"]
            update_output(window)
        elif event == "-after-":
            template["after"] = values["-after-"]
            update_output(window)
        elif event == "-copy-":
            copy_to_clipboard(window, values["output"])
            window["-operation_status-"].update("Copied to clipboard.")
        elif event == "-save-":
            template_name = sg.popup_get_text("Enter a name for the template")
            if template_name:
                filename = os.path.join("template " + template_name + ".json")
                template_files.append(filename)
                template_files.sort()
                template["before"] = values["-before-"]
                template["after"] = values["-after-"]
                with open(filename, "w") as f:
                    f.write(json.dumps(template, indent=4))
                window["-operation_status-"].update("Template saved.")
                window["-template_files-"].update(values=template_files)
        elif event == "-template_files-":
            if values["-template_files-"] != "Default":
                filename = values["-template_files-"]
                with open(filename, "r") as f:
                    loaded = json.loads(f.read())
                window["-before-"].update(value=loaded["before"])
                window["-after-"].update(value=loaded["after"])
                template = loaded
                window["-operation_status-"].update("Template loaded.")
                update_output(window)
                # window.perform_long_operation(lambda: update_output(window), "-thread-"), dont think this is needed
        elif event == "GitHub link" or event == "GitHub link [Ctrl + U]":
            webbrowser.open("https://github.com/MrRazamataz/PremierPro-Markers")
        elif event == "Check for updates":
            window["-program_log-"].update(f"{version_string} (checking for updates...)")
            window.perform_long_operation(lambda: check_for_update(window), "-thread-")
        elif event == "-thread-":
            if values["-thread-"][0]:
                release_notes = values["-thread-"][1]
                sg.Popup(f"Update available!\nRelease notes:\n{release_notes}\nDownload from GitHub (Ctrl + U)!", title="Update")
        elif event == "Settings":
            changed = False
            setting_window = sg.Window("Settings",
                                       [
                                           [sg.Text("Settings", font=("Arial", 20))],
                                           [sg.Text("DPI scaling. This will fix blurryness on scaling other than 100%. If it looks wierd for some reason, you can disable it below.")],
                                           [sg.Check("DPI scaling", default=get_setting_boolean("dpi"), key="-dpi-", enable_events=True)],
                                       ]
            )
            while True:
                setting_event, setting_values = setting_window.read()
                if debug:
                    print(setting_event, setting_values)
                if setting_event == sg.WIN_CLOSED:
                    break
                elif setting_event == "-dpi-":
                    changed = True
                    settings["dpi"] = setting_values["-dpi-"]
                    with open("settings.json", "w") as f:
                        f.write(json.dumps(settings, indent=4))
            if changed:
                sg.popup("You need to restart the program for some changes to take effect.", title="Settings changed detected")


if __name__ == '__main__':
    main()
