import os
from tkinter import filedialog, messagebox
import customtkinter as ctk
import xml.etree.ElementTree as ET
from xml.dom import minidom

global filename
global file_path

def clear_screen(root):
    for widget in root.winfo_children():
        widget.destroy()

global window_width
global window_height
global window_x
global window_y

window_width = None
window_height = None
window_x = None
window_y = None

def go_back(current_screen, create_main_screen_func):
    # Instead of destroying the window, clear its content and rebuild main screen
    clear_screen(current_screen)
    create_main_screen_func()


def toggle_mode():
    current_mode = ctk.get_appearance_mode()
    if current_mode == "Light":
        new_mode = "dark"
    else:
        new_mode = "Light"
    ctk.set_appearance_mode(new_mode)

# valideer time fields
def validateTime(input):
    try:
        parts = input.split(":")
    except:
        return False
    if len(parts) != 3:
        return False
    try:
        hours, minutes, seconds = map(int, parts)
        if 0 <= hours <= 23 and 0 <= minutes <= 59 and 0 <= seconds <= 59:
            return True
    except ValueError:
        pass
    return False

#convert HH:MM:SS naar  x deciseconden
def convertTime(input):
    parts = input.split(":")
    total = 0
    try:
        total = int(parts[0])*60
        total = (total + int(parts[1])) * 60
        total = (total + int(parts[2])) * 10
    except:
        print("fout in omzetten")
    return total

def convertTimeToText(input):
    total_seconds = input // 10  # Convert deciseconds to seconds
    hours, remainder = divmod(total_seconds, 3600)  # Calculate hours
    minutes, seconds = divmod(remainder, 60)  # Calculate minutes and seconds

    # Format the time as 'HH:MM:SS'
    time_format = '{:02d}:{:02d}:{:02d}'.format(hours, minutes, seconds)

    return time_format

def show_info_popup(title, message, root):
    # Create a popup window
    popup = ctk.CTkToplevel()
    popup.title(title)

    # Make the popup transient, so it stays on top of the root window
    popup.transient(root)

    # Update the window to make sure it's been drawn before we try to get its size
    popup.update()

    # Get the window's width and height
    window_width = popup.winfo_width()
    window_height = popup.winfo_height()

    # Get the screen's width and height
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()

    # Calculate the position to center the window
    position_top = (screen_height - window_height) // 2
    position_left = (screen_width - window_width) // 2

    # Position the window at the center of the screen
    popup.geometry("+{}+{}".format(position_left, position_top))

    # Create a label to display the message
    message_label = ctk.CTkLabel(popup, text=message)
    message_label.pack(padx=20, pady=10)

    # Create a button to close the popup
    close_button = ctk.CTkButton(popup, text="Close", command=popup.destroy)
    close_button.pack(pady=10)

def process_sequences(data):
    sequences = []

    for sequence in data:
        # Extract start and end samples
        start_sample = sequence[0][0]
        end_sample = sequence[-1][0]

        # Extract start and end sensors (channels)
        start_sensor = sequence[0][1]
        end_sensor = sequence[-1][1]

        # Create a new sequence dictionary
        seq_dict = {
            "startSample": int(start_sample * 10),
            "endSample": int(end_sample * 10),
            "startChannel": int(start_sensor.split('_')[1]) - 2,
            "endChannel": int(end_sensor.split('_')[1]) - 2,
            "ranges": []
        }

        # Add range data
        for entry in sequence:
            sensor_value = entry[2]
            sample, sensor, _ = entry
            channel = int(sensor.split('_')[1]) -2
            seq_dict["ranges"].append({
                "startSample": int(start_sample * 10),
                "endSample": int(end_sample * 10),
                "channel": channel,
                "maxSample": int(sample * 10),
                "maxValue": int(sensor_value)
            })

        sequences.append(seq_dict)

    return sequences

def sequences_to_xml(sequences, distance_between_sensors):
    root = ET.Element("sequences")

    for seq in sequences:
        time = int((seq["endSample"]) - int(seq["startSample"]))
        if (time == 0):
            velocity = "INF"
            dir = 'Synchronous'
        else:
            velocity = ((int(seq["endChannel"]) - int(seq["startChannel"])) * distance_between_sensors ) / (time / 10)
            if int(velocity) > 0:
                dir = 'Antegrade'
            elif int(velocity) < 0:
                dir = 'Retrograde'
            elif(int(velocity) == 0):
                dir = 'Synchronous'

        seq_elem = ET.SubElement(root, "sequence", {
            "dir": dir,
            "vel": str(velocity),
            "startSample": str(seq["startSample"]),
            "endSample": str(seq["endSample"]),
            "startChannel": str(seq["startChannel"]),
            "endChannel": str(seq["endChannel"])
        })

        for r in seq["ranges"]:
            ET.SubElement(seq_elem, "range", {
                "startSample": str(r["startSample"]),
                "endSample": str(r["endSample"]),
                "channel": str(r["channel"]),
                "maxSample": str(r["maxSample"]),
                "maxValue": str(r["maxValue"])
            })

    xml_str = minidom.parseString(ET.tostring(root, encoding='utf-8')).toprettyxml(indent="    ")
    return xml_str.split('\n', 1)[-1]  # Remove the first line which contains the redundant XML declaration

def write_xml_to_file(xml_output, filename):
    filename = filename.split('.')[0]
    save_path = filedialog.asksaveasfilename(defaultextension=".seq", filetypes=[("Sequences files", "*.seq")], initialfile = f"{filename.split('.seq')[0]}_detected.seq")
    with open(save_path, 'w', encoding='utf-8') as file:
        file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        file.write(xml_output)
        print("XML file 'hrm_output' has been created.")
