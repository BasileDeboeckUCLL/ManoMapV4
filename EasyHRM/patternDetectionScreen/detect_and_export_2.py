import os
from utils import process_sequences
import pandas as pd
from utils import sequences_to_xml, write_xml_to_file, convertTime, validateTime, show_info_popup
from tkinter import filedialog
import numpy as np
from scipy.ndimage import label

global result
result = []

# Read the input file into a DataFrame
global input_file_path
input_file_path = ''

visible_sensors = (1,40)
detection_threshold = 100
zone_threshold = 100
min_pattern_length = 3
maximum_chunk_size = 30
distance_between_sensors = 25

timestamps = pd.DataFrame()
values = pd.DataFrame()
global mask


# Define a custom structure for labeling with diagonal connections
structure = np.array([[1, 1, 1],
                      [1, 1, 1],
                      [1, 1, 1]])

def split_continuous_sensors(patterns):
    continuous_segments = []
    current_segment = []

    for i, pattern in enumerate(patterns):
        if not current_segment:
            current_segment.append(pattern)
        else:
            last_sensor_number = int(current_segment[-1][1].split('_')[1])
            current_sensor_number = int(pattern[1].split('_')[1])
            if current_sensor_number == last_sensor_number + 1:
                current_segment.append(pattern)
            else:
                continuous_segments.append(current_segment)
                current_segment = [pattern]
    if current_segment:
        continuous_segments.append(current_segment)

    filtered_segments = [segment for segment in continuous_segments if len(segment) >= min_pattern_length]
    return filtered_segments

def import_txt_file_detection(file_label, button_export, button_approximate, button_detect_events):
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file_path and os.path.isfile(file_path):
        global filename
        filename = os.path.basename(file_path)
        file_label.configure(text="Selected Text File: " + filename, font=("Arial", 12))
        # button_export.configure(state='normal')
        button_approximate.configure(state='normal')
        button_detect_events.configure(state='normal')
        global input_file_path
        input_file_path = file_path
    else:
        file_label.configure(text="No file selected")
        button_export.configure(state='disabled')
        button_approximate.configure(state='disabled')
        button_detect_events.configure(state='disabled')
        print("No file selected.")
    return input_file_path

def approximate_broken_sensor(broken_sensor_entries):
    # Read the data from the file
    with open(input_file_path, 'r') as file:
        lines = file.readlines()
    
    # Initialize an empty list to hold the processed data
    data = []
    
    # Process each line in the file
    for line in lines:
        if line.strip():  # Skip any empty lines
            parts = line.split()
            time = float(parts[0])  # Convert the first column to float for time
            sensors = list(map(int, parts[1:]))  # Convert the rest to integers for sensor values
            data.append([time] + sensors)
    
    # Convert the list to a numpy array for easier manipulation
    data = np.array(data, dtype=object)
    
    for broken_sensor in broken_sensor_entries:
        if not broken_sensor.get().strip(' ') == '':
            broken_sensor_index = int(broken_sensor.get())
            # Replace the broken sensor values with the average of the previous and next sensor values
            for row in data:
                if broken_sensor_index == 1:
                    row[broken_sensor_index+1] = row[broken_sensor_index+2]
                elif broken_sensor_index == len(row) - 2:
                    row[broken_sensor_index+1] = row[broken_sensor_index]
                else:
                    row[broken_sensor_index+1] = int(round((row[broken_sensor_index] + row[broken_sensor_index+2]) / 2))

    # Prompt the user to select where to save the new file
    save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], initialfile = f"{filename.split('.txt')[0]}_approximated.txt")
    
    # Save the modified data back to the file in the same format
    with open(save_path, 'w') as file:
        for row in data:
            # Format row[0] and strip '.0' if necessary
            row_0_str = f"{row[0]:.1f}".rstrip('0').rstrip('.') if row[0] % 1 == 0 else f"{row[0]:.1f}"
            formatted_row = f"{row_0_str}\t\t {row[1]:d}"
            for value in row[2:]:
                formatted_row += f"\t {value:d}"
            file.write(formatted_row + '\n')

    print(f"File saved as: {save_path}")

def read_data(total_seconds):
    # Read the data into a DataFrame
    global dataframe

    dataframe = pd.read_csv(input_file_path, sep=" ", header=None)

    dataframe = dataframe[dataframe.iloc[:, 0] > total_seconds]

    # Remove milliseconds from timestamps
    dataframe = dataframe[dataframe[0] == dataframe[0].astype(int)]

    # Separate the timestamps and the values
    global timestamps
    timestamps = dataframe.iloc[:, 0]
    global values
    values = dataframe.iloc[:, 1:]

    # Keep only the visible sensors
    values = values.iloc[:, int(round(visible_sensors[0]))-1:int(round(visible_sensors[1]))]

    # Create a mask where values are greater than the threshold
    global mask
    mask = values > zone_threshold


def find_patterns(zone_df):
    patterns = []
    for sensor_id, group in zone_df.groupby('Sensor_ID'):
        max_sensor_value = group['Value'].max()
        sensor_timestamp = group.loc[group['Value'].idxmax(), 'Timestamp']
        if max_sensor_value > detection_threshold:
            pattern = (sensor_timestamp, f'sensor_{sensor_id}', max_sensor_value)
            patterns.append(pattern)
    return patterns

def define_chunks_and_get_patterns():
    # Use the label function to find connected regions
    labeled_array, num_features = label(mask, structure)

    # Extract the zones, their values, timestamps, and sensor IDs
    results = {}
    for zone in range(1, num_features + 1):
        zone_indices = np.argwhere(labeled_array == zone)
        zone_data = [(timestamps.iloc[i], j + 1, values.iloc[i, j]) for i, j in zone_indices]  # j + 1 to get the sensor_id starting from 1
        results[zone] = zone_data
    
    global result
    result = []
    # Detect patterns in each zone
    for zone, data in results.items():
        zone_df = pd.DataFrame(data, columns=['Timestamp', 'Sensor_ID', 'Value'])
        zone_patterns = find_patterns(zone_df)
        split_zone_patterns = split_continuous_sensors(zone_patterns)

        if split_zone_patterns:
            for pattern in split_zone_patterns:
                if len(pattern) >= min_pattern_length:
                    result.append(pattern)
    return result

def compute_patterns(sliders, advanced_sliders, time_entries, settings_frame, button_export):
    global detection_threshold
    detection_threshold = int(round(advanced_sliders[0].get()))

    global min_pattern_length
    min_pattern_length = int(round(advanced_sliders[1].get()))
    
    global distance_between_sensors
    distance_between_sensors = int(round(advanced_sliders[2].get()))

    global zone_threshold
    zone_threshold = int(round(advanced_sliders[3].get()))

    global visible_sensors
    visible_sensors = sliders[0].get()

    # Extract time from entries
    hour = time_entries[0].get() or 0
    minute = time_entries[1].get() or 0
    second = time_entries[2].get() or 0
    time_string = f"{hour}:{minute}:{second}"

    # Validate and convert time
    if validateTime(time_string):
        total_seconds = round(convertTime(time_string) / 10)
    else:
        print("Invalid time format")
    
    read_data(total_seconds)

    global result
    result = define_chunks_and_get_patterns()
    show_info_popup("Succes", "Detection Completed", settings_frame)

    #Enable export button after detection
    button_export.configure(state='normal')

def exportToXML():
    sequences = process_sequences(result)
    xml_output = sequences_to_xml(sequences, distance_between_sensors)
    write_xml_to_file(xml_output, filename)