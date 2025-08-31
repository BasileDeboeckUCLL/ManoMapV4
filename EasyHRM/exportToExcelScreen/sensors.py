import customtkinter as ctk
from CTkRangeSlider import *
from exportToExcelScreen.export import remove_disabled_sections, add_disabled_sections, reset_disabled_sections

def create_sensors_frame(root):
    sensors_frame = ctk.CTkFrame(root)
    sensors_frame.pack(pady=10, padx=10, fill="both", expand=True)

    # Format for each setting: (label_text, from_, to, start_value, end_value)
    colonregions = [
        ("Ascending:", 1, 80, 1, 62),
        ("Transverse:", 1, 80, 63, 68),
        ("Descending:", 1, 80, 69, 72),
        ("Sigmoid:", 1, 80, 73, 76),
        ("Rectum:", 1, 80, 77, 80)
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
        value_labels[i][0].configure(text=f"{start}")
        value_labels[i][1].configure(text=f"{end}")

        # Update the next slider's start value
        if i < len(sliders) - 1:
            sliders[i + 1].set([end + 1, sliders[i + 1].get()[1]])
            value_labels[i + 1][0].configure(text=f"{end + 1}")
            value_labels[i + 1][1].configure(text=f"{int(sliders[i + 1].get()[1])}")

        # Update the previous slider's end value
        if i > 0:
            prev_start, prev_end = sliders[i - 1].get()
            sliders[i - 1].set([prev_start, start - 1])
            value_labels[i - 1][0].configure(text=f"{int(prev_start)}")
            value_labels[i - 1][1].configure(text=f"{start - 1}")
    
    def checkbox_event(i, label_text):
        stripped_label_text = label_text.strip(":")
        if checkboxes[i].get() == "on":
            sliders[i].configure(state="normal", progress_color="grey", button_color="#1F6AA5")
            remove_disabled_sections(stripped_label_text)
            # Adjust the previous and next sliders
            if i > 0:
                prev_start = sliders[i - 1].get()[0]
                sliders[i - 1].set([prev_start, sliders[i].get()[0] - 1])
                value_labels[i - 1][1].configure(text=f"{int(round(sliders[i].get()[0] - 1))}")

            if i < len(sliders) - 1:
                next_end = sliders[i + 1].get()[1]
                sliders[i + 1].set([sliders[i].get()[1] + 1, next_end])
                value_labels[i + 1][0].configure(text=f"{int(round(sliders[i].get()[1] + 1))}")

            # Ensure all previous checkboxes are checked
            for j in range(i+1, len(checkboxes)):
                checkboxes[j].select()
                checkbox_event(j, colonregions[j][0])

        else:
            sliders[i].configure(state="disabled",  progress_color="transparent", button_color="grey")
            add_disabled_sections(stripped_label_text)
            if i < len(sliders) - 1:
                start_value = sliders[i].get()[0]
                sliders[i + 1].set([start_value, sliders[i + 1].get()[1]])
                value_labels[i + 1][0].configure(text=f"{int(round(start_value))}")

            # Ensure all previous checkboxes are unchecked
            for j in range(i-1, -1, -1):
                checkboxes[j].deselect()
                checkbox_event(j, colonregions[j][0])

    checkboxes = []
    for i, (label_text, from_, to, start_value, end_value) in enumerate(colonregions):
        setting_checkbox = ctk.CTkCheckBox(sensors_frame, text=label_text, onvalue="on", offvalue="off", command=lambda i=i, label_text=label_text: checkbox_event(i, label_text))
        setting_checkbox.grid(row=i, column=0, padx=5, pady=5)
        setting_checkbox.select()
        checkboxes.append(setting_checkbox)

        # label = ctk.CTkLabel(sensors_frame, text=label_text)
        # label.grid(row=i, column=0, padx=5, pady=5)

        # Value label to the left of the slider
        value_label1 = ctk.CTkLabel(sensors_frame, text="")
        value_label1.grid(row=i, column=1, padx=5, pady=5)

        # Slider
        slider = CTkRangeSlider(sensors_frame, from_=from_, to=to, command=lambda value, i=i: update_value_label(value, i))
        slider.grid(row=i, column=2, padx=5, pady=5, sticky="ew")
        slider.set([start_value, end_value])
        sliders.append(slider)

        # Value label to the right of the slider
        value_label2 = ctk.CTkLabel(sensors_frame, text="")
        value_label2.grid(row=i, column=3, padx=5, pady=5)

        # Append the tuple of value labels
        value_labels.append((value_label1, value_label2))

        # Call the update_value_label function with the initial values of the slider
        update_value_label((slider.get()[0], slider.get()[1]), i)

    settings_sliders = []
    for i, (label_text, from_, to, default_value) in enumerate(settings):
        row_index = i + len(colonregions)
        label = ctk.CTkLabel(sensors_frame, text=label_text)
        label.grid(row=row_index, column=0, padx=5, pady=5)

        value = ctk.IntVar()
        value.set(default_value)
        slider = ctk.CTkSlider(sensors_frame, from_=from_, to=to, variable=value)
        slider.grid(row=row_index, column=2, padx=5, pady=5, sticky="ew")
        
        settings_sliders.append(slider)

        value_label = ctk.CTkEntry(sensors_frame, textvariable=value, width=40)
        value_label.grid(row=row_index, column=1, padx=5, pady=5)

    return sliders, settings_sliders
