import sys
import traceback
import urllib.request
import webbrowser
from importlib.metadata import version
import csv
import os
import json
import ctypes
import platform
import time



# check that is version 4.60.4 or below NOT 5.X.X or higher
if int(version('PySimpleGUI').split(".")[0]) >= 5:
    print("The version of PySimpleGUI is too high (up to 4.60.4 is supported) because it now requires registration or payment. Automatically downloading 4.60.4...")
    import requests
    url = "https://raw.githubusercontent.com/andor-pierdelacabeza/PySimpleGUI-4-foss/foss/PySimpleGUI.py"
    r = requests.get(url)
    with open("PySimpleGUI.py", "wb") as f:
        f.write(r.content)
    print("Downloaded PySimpleGUI 4.60.4. Program will now start.")
import PySimpleGUI as sg


version = 9
version_string = "v1.0.0"
debug = False
template_filename = "template.json"
template = {}
markers = []
blink = {}
output = ""

if not os.path.isfile("settings.json"):
    with open("settings.json", "w") as f:
        f.write(json.dumps({"dpi": True}, indent=4))
        settings = {"dpi": True}
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
# check if windows or linux
if platform.system() == "Windows":
    if not os.path.isfile("run.bat"):
        with open("run.bat", "w") as f:
            f.write("""@echo off
echo You can safely close this window.
start pythonw main.py
exit
""")
            print("Created run.bat")

# find other filesnames that contain "template" and ".json"
template_files = [f for f in os.listdir() if "template" in f and ".json" in f and f != template_filename]
if len(template_files) > 0:
    template_files.sort()
else:
    template_files = ["Last used"]


def check_for_update(window: sg.Window):
    global debug, last_error
    print("Checking for updates...")
    window.set_title("PremierPro Markers")
    try:
        with urllib.request.urlopen(
                "https://raw.githubusercontent.com/MrRazamataz/PremierPro-Markers/master/version.json") as url:
            data = json.loads(url.read().decode())
            if data["version"] > version:
                print("Update available!")
                release_notes = data["release_notes"]
                window["-program_log-"].update(f"{version_string} (outdated - update available)")
                window.set_title(window.Title + " (outdated)")
                return True, release_notes
            elif data["version"] == version:
                print("Up to date!")
                window["-program_log-"].update(version_string)
            elif data["version"] < version:
                print("You're using a newer version than the latest release, DEBUG MODE ON!")
                print("Make sure to PR any cool changes you make :)\n----")
                debug = True
                window["-program_log-"].update(f"{version_string} (modified)")
                update_nav(window, "debug")
                window.set_title(window.Title + " (DEBUG MODE)")
    except Exception as e:
        print("Error checking for updates: " + str(e))
        print(traceback.format_exc())
        show_error_window(e)
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
    global last_error
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
        show_error_window(e)
        return False


def update_output(window: sg.Window):
    global output
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


def restart():
    os.execv(sys.executable, ['python'] + sys.argv)


def update_nav(window: sg.Window, layout: str or None):
    nav_menu_def = [['File', ["Open... [Ctrl + O]", "Save template [Ctrl + S]", "Exit [Ctrl + W]"]],
                    ['Edit', ["Copy output [Ctrl + C]", "Settings"]],
                    ['Templates', template_files],
                    ['Help', ["GitHub link [Ctrl + U]", "Check for updates"]]]
    if layout == "debug":
        nav_menu_def.append(['Debug', ["Restart", "Variables", "Run GC"]])
    window['-menu-'](nav_menu_def)


def update_templates(window: sg.Window):
    template_files = [f for f in os.listdir() if "template" in f and ".json" in f and f != template_filename]
    if len(template_files) > 0:
        template_files.sort()
    else:
        template_files = ["Last used"]
    #window["-template_files-"].update(values=template_files)
    nav_menu_def = window["-menu-"].MenuDefinition
    nav_menu_def[2][1] = template_files
    window['-menu-'](nav_menu_def)


def show_error_window(error, buggered=False):
    error_index = {
        "UTF-16 stream does not start with BOM": "Make sure the CSV file was created with PremierPro!",
        "[Errno 2] No such file or directory": "The template file cannot be found, did you delete it?"
    }
    ignore_error_index = ["[Errno 22] Invalid argument"]
    if str(error) in ignore_error_index:
        return
    message = error_index.get(str(error), "No known solution found.")
    error_window = sg.Window("Error", [
        [sg.Text(message, font=20)],
        [sg.Text("Error message:", font=10)],
        [sg.Text(error, font=("Consolas", 10))],
        [sg.Text("This error is unrecoverable and the app will now restart, sorry.", key="-buggered-", visible=False, text_color="red", background_color="black")],
        [sg.Text("The program has attempted to save the last output", key="-attempt-", visible=False, text_color="green", background_color="black")],
        [sg.Button("Close", key="-close-"), sg.Button("Copy error", key="-copy_error-")]
    ], disable_minimize=True, keep_on_top=True, finalize=True)
    if buggered:
        error_window["-buggered-"].update(visible=True)
        error_window["-attempt-"].update(visible=True)
    while True:
        event, values = error_window.read()
        if event == "-close-" or event == sg.WIN_CLOSED:
            error_window.close()
            if buggered:
                restart()
            break
        elif event == "-copy_error-":
            copy_to_clipboard(error_window, str(error))


def custom_exception_handler(exc_type, exc_value, exc_traceback):
    global output
    # attempt to save all data whilst withholding the crash
    with open("saved_output.txt", "w") as file:
        file.write(output)
    show_error_window(exc_value, True)


sys.excepthook = custom_exception_handler


def main():
    global template, debug, output
    make_dpi_aware()
    markers_layout = [
        [sg.Text(
            'Automagically create YouTube chapters from PremierPro markers. Made by MrRazamataz - inspired by RavinMaddHatter.',
            expand_x=True, key="-info_text-")],
        [sg.FileBrowse("Browse", key="-file_browse-", enable_events=True,
                       file_types=(("CSV File", "*.csv"),))],
        [sg.Text("Output:")],
        [sg.Multiline(key='output', expand_x=True, expand_y=True, right_click_menu=["&Right", ["Copy"]], enable_events=True)],
        [sg.Button("Copy", key="-copy-", tooltip="Copy to clipboard [Ctrl + C]."), sg.Text("First timestamp modified to be 00:00!", text_color="red", key="-modified_warning-", background_color="black", visible=False)],
    ]
    template_layout = [
        [sg.Text('Create the template for your description generator.', expand_x=True), sg.Push(),
         sg.DropDown(template_files, key="-template_files-", enable_events=True, default_value="Last used",
                     readonly=True, auto_size_text=True, size=(20, 1))],
        [sg.Text("Before chapters:")],
        [sg.Multiline(key='-before-', expand_x=True, expand_y=True, default_text=template["before"], enable_events=True,
                      size=(20, 10), right_click_menu=["&Right", ["Paste to before chapters"]])],
        [sg.Text("After chapters:")],
        [sg.Multiline(key='-after-', expand_x=True, expand_y=True, default_text=template["after"], enable_events=True,
                      size=(20, 10), right_click_menu=["&Right", ["Paste to after chapters"]])],
        [sg.Text(
            "The defaut template is the last used template. If you want to save a template, enter a name in the box below and click save.\nYou can then switch between saved templates using the dropdown above.",
            size=(90, 3))],
        [sg.Button("Save as template", key="-save-", tooltip="Save template.")],
    ]

    logs_layout = [
        [sg.Output(size=(70, 20), key='output', expand_x=True, expand_y=True)],
    ]
    layout = [
        [sg.Menu(["File"], key="-menu-")],
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
    update_nav(window, None)
    if os.path.isfile("saved_output.txt"):
        with open("saved_output.txt", "r") as file:
            output = file.read()
            window["output"].update(output)
            window["-operation_status-"].update("Recovered output from crash.")
        os.remove("saved_output.txt")
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
                    window["-operation_status-"].update("Error opening file")
                update_output(window)
        elif event == '-file_browse-':
            window["-operation_status-"].update("Select a file...")
            if values['-file_browse-']:
                if open_csv_file(values['-file_browse-']) == False:
                    window["-operation_status-"].update("Error opening file")
                update_output(window)
        elif event == "-before-":
            template["before"] = values["-before-"]
            update_output(window)
        elif event == "-after-":
            template["after"] = values["-after-"]
            update_output(window)
        elif event == "-copy-" or event == "Copy":
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
                update_templates(window)
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
                                           [sg.Button("Force Enable Debug Mode", key="-debug-", tooltip="This will enable debug mode even if the program is not newer than release.")],
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
                elif setting_event == "-debug-":
                    debug = True
                    update_nav(window, "debug")
                    window.set_title(window.Title + " (DEBUG MODE)")
                    setting_window.close()
            if changed:
                if sg.PopupYesNo("You need to restart the program for some changes to take effect.\nDo you wish to restart now?", title="Settings changed detected") == "Yes":
                    restart()

        elif event == "Paste to before chapters":
            window['-before-'].update(window.TKroot.clipboard_get())
        elif event == "Paste to after chapters":
            window['-after-'].update(window.TKroot.clipboard_get())
        elif event == "Restart":
            restart()
        elif event == "Variables":
            # popup window with the current variables
            # format them nicely with new lines
            variables = locals()
            default_text = ''.join([f"{k}: {v}\n" for k, v in variables.items()])
            menu_bar_list = window["-menu-"].MenuDefinition
            default_text += f"\n---------------\nPremier Markers Specifics:\nMenu bar list:\n{menu_bar_list}\n"
            variableWindow = sg.Window("Variables", [
                [sg.Multiline(default_text=default_text, expand_x=True, expand_y=True, disabled=True)]
            ], size=(800, 600), resizable=True)
            variableWindow.read()
            variableWindow.close()
        elif "template" and ".json" in event:
            filename = event
            try:
                with open(filename, "r") as f:
                    loaded = json.loads(f.read())
                window["-before-"].update(value=loaded["before"])
                window["-after-"].update(value=loaded["after"])
                template = loaded
                window["-operation_status-"].update("Template loaded.")
                update_output(window)
            except Exception as e:
                show_error_window(e)
        elif event == "Run GC":
            import gc
            gc.collect()
        elif event == "output":
            output = values["output"]


if __name__ == '__main__':
    main()
