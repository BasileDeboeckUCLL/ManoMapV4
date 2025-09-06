import customtkinter as ctk
from CTkRangeSlider import *
from exportToExcelScreen.export import remove_disabled_sections, add_disabled_sections, reset_disabled_sections

def create_sensors_frame(root):
    sensors_frame = ctk.CTkFrame(root)
    sensors_frame.pack(pady=10, padx=10, fill="both", expand=True)

    # Move sensors label to top
    sensors_label = ctk.CTkLabel(sensors_frame, text="Sensors", font=("Arial", 14, "bold"))
    sensors_label.grid(row=0, column=0, columnspan=4, pady=(10, 5))

    # Format for each setting: (label_text, from_, to, start_value, end_value)
    colonregions = [
        ("Ascending:", 1, 80, 1, 16),
        ("Transverse:", 1, 80, 17, 32),
        ("Descending:", 1, 80, 33, 48),
        ("Sigmoid:", 1, 80, 49, 64),
        ("Rectum:", 1, 80, 65, 80)
    ]
    sliders = []
    value_labels = []

    reset_disabled_sections()

    # Format for each setting: (label_text, from_, to, default_value)
    settings = [
        ("Distance between sensors (mm)", 1, 200, 25),
    ]

    def update_value_label(value, i):
        start, end = int(round(value[0], 0)), int(round(value[1], 0))
        
        # Calculate bounds based on adjacent sliders
        min_start = 1  # Absolute minimum
        max_end = 80   # Absolute maximum
        
        # Get bounds from previous slider (if exists and enabled)
        if i > 0 and checkboxes[i-1].get() == "on":
            prev_end = int(round(sliders[i-1].get()[1], 0))
            min_start = prev_end + 1
        
        # Get bounds from next slider (if exists and enabled)
        if i < len(sliders) - 1 and checkboxes[i+1].get() == "on":
            next_start = int(round(sliders[i+1].get()[0], 0))
            max_end = next_start - 1
        
        # Constrain the values to prevent overlapping
        start = max(min_start, start)
        end = min(max_end, end)
        
        # Ensure start <= end with minimum gap of 1
        if start >= end:
            if start == min_start:
                end = start + 1
                end = min(end, max_end)
            elif end == max_end:
                start = end - 1
                start = max(start, min_start)
            else:
                # Adjust based on which boundary was hit
                mid = (min_start + max_end) // 2
                if start <= mid:
                    end = start + 1
                    end = min(end, max_end)
                else:
                    start = end - 1
                    start = max(start, min_start)
        
        # Update the current slider with constrained values
        if (start, end) != (int(round(value[0], 0)), int(round(value[1], 0))):
            sliders[i].set([start, end])
        
        # Update labels
        value_labels[i][0].configure(text=f"{start}")
        value_labels[i][1].configure(text=f"{end}")
        
        # Update adjacent sliders if needed
        update_adjacent_sliders(i, start, end)

    def update_adjacent_sliders(current_i, current_start, current_end):
        # Update next slider if it exists and is enabled
        if current_i < len(sliders) - 1 and checkboxes[current_i + 1].get() == "on":
            next_start, next_end = sliders[current_i + 1].get()
            new_next_start = max(current_end + 1, int(round(next_start, 0)))
            
            if new_next_start != int(round(next_start, 0)):
                # Ensure the next slider's end doesn't get squeezed too much
                new_next_end = max(new_next_start + 1, int(round(next_end, 0)))
                sliders[current_i + 1].set([new_next_start, new_next_end])
                value_labels[current_i + 1][0].configure(text=f"{new_next_start}")
                value_labels[current_i + 1][1].configure(text=f"{new_next_end}")
                
                # Recursively update subsequent sliders
                update_adjacent_sliders(current_i + 1, new_next_start, new_next_end)
        
        # Update previous slider if it exists and is enabled
        if current_i > 0 and checkboxes[current_i - 1].get() == "on":
            prev_start, prev_end = sliders[current_i - 1].get()
            new_prev_end = min(current_start - 1, int(round(prev_end, 0)))
            
            if new_prev_end != int(round(prev_end, 0)):
                # Ensure the previous slider's start doesn't get squeezed too much
                new_prev_start = min(new_prev_end - 1, int(round(prev_start, 0)))
                sliders[current_i - 1].set([new_prev_start, new_prev_end])
                value_labels[current_i - 1][0].configure(text=f"{new_prev_start}")
                value_labels[current_i - 1][1].configure(text=f"{new_prev_end}")
                
                # Recursively update previous sliders
                update_adjacent_sliders(current_i - 1, new_prev_start, new_prev_end)
    
    def checkbox_event(i, label_text):
        stripped_label_text = label_text.strip(":")
        if checkboxes[i].get() == "on":
            sliders[i].configure(state="normal", progress_color="grey", button_color="#1F6AA5")
            remove_disabled_sections(stripped_label_text)
            
            # Recalculate positions for all sliders to prevent overlaps
            recalculate_all_slider_positions()
            
            # Ensure all following checkboxes are checked
            for j in range(i+1, len(checkboxes)):
                if checkboxes[j].get() == "off":
                    checkboxes[j].select()
                    checkbox_event(j, colonregions[j][0])

        else:
            sliders[i].configure(state="disabled", progress_color="transparent", button_color="grey")
            add_disabled_sections(stripped_label_text)
            
            # Ensure all previous checkboxes are unchecked
            for j in range(i-1, -1, -1):
                if checkboxes[j].get() == "on":
                    checkboxes[j].deselect()
                    checkbox_event(j, colonregions[j][0])
            
            # Recalculate positions after disabling
            recalculate_all_slider_positions()

    def recalculate_all_slider_positions():
        """Recalculate positions for all enabled sliders to prevent overlaps"""
        enabled_indices = [i for i in range(len(checkboxes)) if checkboxes[i].get() == "on"]
        
        for idx, i in enumerate(enabled_indices):
            current_start, current_end = sliders[i].get()
            
            # Calculate minimum start based on previous enabled slider
            min_start = 1
            if idx > 0:
                prev_i = enabled_indices[idx - 1]
                prev_end = int(round(sliders[prev_i].get()[1], 0))
                min_start = prev_end + 1
            
            # Calculate maximum end based on next enabled slider
            max_end = 80
            if idx < len(enabled_indices) - 1:
                next_i = enabled_indices[idx + 1]
                next_start = int(round(sliders[next_i].get()[0], 0))
                max_end = next_start - 1
            
            # Adjust current slider if needed
            new_start = max(min_start, int(round(current_start, 0)))
            new_end = min(max_end, int(round(current_end, 0)))
            
            # Ensure minimum width
            if new_end <= new_start:
                if max_end - min_start >= 1:
                    new_end = new_start + 1
                    if new_end > max_end:
                        new_end = max_end
                        new_start = new_end - 1
                else:
                    # Not enough space, adjust neighboring sliders
                    new_start = min_start
                    new_end = max_end
            
            if (new_start, new_end) != (int(round(current_start, 0)), int(round(current_end, 0))):
                sliders[i].set([new_start, new_end])
                value_labels[i][0].configure(text=f"{new_start}")
                value_labels[i][1].configure(text=f"{new_end}")

    def reset_sensors():
        """Reset all sensors to their default values"""
        # Default values from the original colonregions list
        default_values = [(1, 16), (17, 32), (33, 48), (49, 64), (65, 80)]
        
        # Reset all checkboxes to checked
        for i, checkbox in enumerate(checkboxes):
            checkbox.select()
        
        # Reset disabled sections
        reset_disabled_sections()
        
        # Reset all sliders to default values
        for i, (start_val, end_val) in enumerate(default_values):
            sliders[i].set([start_val, end_val])
            sliders[i].configure(state="normal", progress_color="grey", button_color="#1F6AA5")
            value_labels[i][0].configure(text=f"{start_val}")
            value_labels[i][1].configure(text=f"{end_val}")
        
        # Reset distance between sensors slider
        settings_sliders[0].set(25)  # Default value for distance between sensors

        # Reset pattern parameters
        pattern_params['long_sensors'].delete(0, 'end')
        pattern_params['long_sensors'].insert(0, "5")
        pattern_params['long_distance'].delete(0, 'end')
        pattern_params['long_distance'].insert(0, "100")
        pattern_params['hapc_sensors'].delete(0, 'end')
        pattern_params['hapc_sensors'].insert(0, "5")
        pattern_params['hapc_distance'].delete(0, 'end')
        pattern_params['hapc_distance'].insert(0, "100")
        pattern_params['hapc_consecutive'].delete(0, 'end')
        pattern_params['hapc_consecutive'].insert(0, "3")
        pattern_params['hapc_amplitude'].delete(0, 'end')
        pattern_params['hapc_amplitude'].insert(0, "100")

    checkboxes = []
    for i, (label_text, from_, to, start_value, end_value) in enumerate(colonregions):
        setting_checkbox = ctk.CTkCheckBox(sensors_frame, text=label_text, onvalue="on", offvalue="off", command=lambda i=i, label_text=label_text: checkbox_event(i, label_text))
        setting_checkbox.grid(row=i+1, column=0, padx=5, pady=5)
        setting_checkbox.select()
        checkboxes.append(setting_checkbox)

        # Value label to the left of the slider
        value_label1 = ctk.CTkLabel(sensors_frame, text="")
        value_label1.grid(row=i+1, column=1, padx=5, pady=5)

        # Slider
        slider = CTkRangeSlider(sensors_frame, from_=from_, to=to, command=lambda value, i=i: update_value_label(value, i))
        slider.grid(row=i+1, column=2, padx=5, pady=5, sticky="ew")
        slider.set([start_value, end_value])
        sliders.append(slider)

        # Value label to the right of the slider
        value_label2 = ctk.CTkLabel(sensors_frame, text="")
        value_label2.grid(row=i+1, column=3, padx=5, pady=5)

        # Append the tuple of value labels
        value_labels.append((value_label1, value_label2))

        # Call the update_value_label function with the initial values of the slider
        update_value_label((slider.get()[0], slider.get()[1]), i)

    settings_sliders = []
    for i, (label_text, from_, to, default_value) in enumerate(settings):
        row_index = i + len(colonregions) + 1
        label = ctk.CTkLabel(sensors_frame, text=label_text)
        label.grid(row=row_index, column=0, padx=5, pady=5)

        value = ctk.IntVar()
        value.set(default_value)
        slider = ctk.CTkSlider(sensors_frame, from_=from_, to=to, variable=value)
        slider.grid(row=row_index, column=2, padx=5, pady=5, sticky="ew")
        
        settings_sliders.append(slider)

        value_label = ctk.CTkEntry(sensors_frame, textvariable=value, width=40)
        value_label.grid(row=row_index, column=1, padx=5, pady=5)

    # Reset button
    reset_button = ctk.CTkButton(sensors_frame, text="Reset Sensors", command=reset_sensors)
    reset_button.grid(row=len(colonregions) + len(settings) + 2, column=0, columnspan=4, padx=5, pady=10, sticky="ew")

    # Pattern Classification Parameters
    pattern_row_start = len(colonregions) + len(settings) + 3
    
    # Long pattern parameters
    long_frame = ctk.CTkFrame(sensors_frame, fg_color="transparent")
    long_frame.grid(row=pattern_row_start, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
    
    ctk.CTkLabel(long_frame, text="Long = min").pack(side="left", padx=2)
    long_sensors_entry = ctk.CTkEntry(long_frame, width=40)
    long_sensors_entry.pack(side="left", padx=2)
    long_sensors_entry.insert(0, "5")
    
    ctk.CTkLabel(long_frame, text="sensors and").pack(side="left", padx=2)
    long_distance_entry = ctk.CTkEntry(long_frame, width=40)
    long_distance_entry.pack(side="left", padx=2)
    long_distance_entry.insert(0, "100")
    
    ctk.CTkLabel(long_frame, text="mm long").pack(side="left", padx=2)
    
    # HA(R)PC parameters
    hapc_frame = ctk.CTkFrame(sensors_frame, fg_color="transparent")
    hapc_frame.grid(row=pattern_row_start + 1, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
    
    ctk.CTkLabel(hapc_frame, text="HA(R)PCs = min").pack(side="left", padx=2)
    hapc_sensors_entry = ctk.CTkEntry(hapc_frame, width=40)
    hapc_sensors_entry.pack(side="left", padx=2)
    hapc_sensors_entry.insert(0, "5")
    
    ctk.CTkLabel(hapc_frame, text="sensors and").pack(side="left", padx=2)
    hapc_distance_entry = ctk.CTkEntry(hapc_frame, width=40)
    hapc_distance_entry.pack(side="left", padx=2)
    hapc_distance_entry.insert(0, "100")
    
    ctk.CTkLabel(hapc_frame, text="mm long of which").pack(side="left", padx=2)
    hapc_consecutive_entry = ctk.CTkEntry(hapc_frame, width=40)
    hapc_consecutive_entry.pack(side="left", padx=2)
    hapc_consecutive_entry.insert(0, "3")
    
    ctk.CTkLabel(hapc_frame, text="is above").pack(side="left", padx=2)
    hapc_amplitude_entry = ctk.CTkEntry(hapc_frame, width=40)
    hapc_amplitude_entry.pack(side="left", padx=2)
    hapc_amplitude_entry.insert(0, "100")
    
    ctk.CTkLabel(hapc_frame, text="mmHg").pack(side="left", padx=2)
    
    # Store pattern parameter entries
    pattern_params = {
        'long_sensors': long_sensors_entry,
        'long_distance': long_distance_entry,
        'hapc_sensors': hapc_sensors_entry,
        'hapc_distance': hapc_distance_entry,
        'hapc_consecutive': hapc_consecutive_entry,
        'hapc_amplitude': hapc_amplitude_entry
    }

    return sliders, settings_sliders, pattern_params