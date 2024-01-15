import traceback
import urllib.request
import webbrowser

import PySimpleGUI as sg
import csv
import os
import json
import ctypes
import platform

version = 2
version_string = "v0.0.2"
debug = False
template_filename = "template.json"
template = {}
markers = []
if not os.path.isfile("settings.json"):
    with open("settings.json", "w") as f:
        f.write(json.dumps({"dpi": True, "version": version}, indent=4))
else:
    with open("settings.json", "r") as f:
        settings = json.loads(f.read())
        if settings["version"] > version:
            version_string = version_string + " (modified)"
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


def check_for_update():
    print("Checking for updates...")
    try:
        with urllib.request.urlopen(
                "https://raw.githubusercontent.com/MrRazamataz/PremierPro-Markers/main/version.json") as url:
            data = json.loads(url.read().decode())
            if data["version"] > version:
                print("Update available!")
                release_notes = data["release_notes"]
                sg.Popup(f"Update available!\nRelease notes:\n{release_notes}\nDownload from GitHub (Ctrl + U)!", title="Update")

    except Exception as e:
        print("Error checking for updates: " + str(e))
        print(traceback.format_exc())


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
    try:
        with open(file_path, encoding='utf-16') as csvfile:
            reader = csv.DictReader(csvfile, delimiter='\t')
            print("File opened successfully")
            markers.clear()
            for i in reader:
                time = format_time(i['In'])
                markers.append(f"{time} {i['Marker Name']}")
            return True
    except Exception as e:
        print("Error opening file: " + str(e))
        print(traceback.format_exc())
        return False


def update_output(window: sg.Window):
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


def main():
    global template
    make_dpi_aware()
    nav_menu_def = [['File', ["Open... [Ctrl + O]", 'Exit']],
                    ['Edit', ["Settings"]],
                    ['Help', ["GitHub link", "Check for updates"]]]

    markers_layout = [
        [sg.Text(
            'Automagically create YouTube chapters from PremierPro markers. Made by MrRazamataz - inspired by RavinMaddHatter.',
            expand_x=True, key="-info_text-")],
        [sg.FileBrowse("Browse", key="-file_browse-", enable_events=True,
                       file_types=(("CSV File", "*.csv"),))],
        [sg.Text("Output:")],
        [sg.Multiline(key='output', expand_x=True, expand_y=True)],
        [sg.Button("Copy", key="-copy-", tooltip="Copy to clipboard.")],
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
        [sg.Menu(nav_menu_def, )],
        [sg.TabGroup(
            [[sg.Tab('Markers', markers_layout), sg.Tab('Template', template_layout), sg.Tab('Log', logs_layout)]],
            expand_x=True, expand_y=True)],

        [sg.StatusBar("To get started, open a file", key="-operation_status-", expand_x=True, auto_size_text=False,
                      right_click_menu=["&Right", ["Open...", "Exit"]], tooltip="The status of the current operation."),
         sg.StatusBar(version_string, key="-program_log-", justification="right", auto_size_text=False,
                      tooltip="Program log.")],
    ]

    window = sg.Window('PremierPro Markers', layout, resizable=True, finalize=True)
    window.bind("<Control-o>", "Open...")
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
        elif event == "GitHub link":
            webbrowser.open("https://github.com/MrRazamataz/PremierPro-Markers")
        elif event == "Check for updates":
            sg.Popup("This feature is not yet implemented. Please manually check for updates on GitHub.",
                     title="Not implemented")


if __name__ == '__main__':
    main()
