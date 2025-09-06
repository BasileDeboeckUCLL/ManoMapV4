import customtkinter as ctk
from exportToExcelScreen.export import exportToXlsx
from utils import clear_screen
from exportToExcelScreen.events import create_event_interface, show_comments
from exportToExcelScreen.sensors import create_sensors_frame
from exportToExcelScreen.importFile import select_input_file


def export_to_excel_screen(root, go_back_func, create_main_screen_func):
    clear_screen(root)

    # Create main frame for layout
    main_frame = ctk.CTkFrame(root)
    main_frame.pack(pady=20, padx=20, fill="both", expand=True)

    # Title Label
    title_label = ctk.CTkLabel(main_frame, text="Data Analysis", font=("Arial", 20, "bold"))
    title_label.grid(row=0, column=0, columnspan=3, pady=10)

    df = None
    file_name = None

    def select_file_and_update_label():
        nonlocal df, file_name
        df, file_name = select_input_file(root, file_label, button_export)

    # Top Buttons
    button_select_input = ctk.CTkButton(main_frame, text="Select Input File", command=lambda: select_file_and_update_label())
    button_select_input.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    file_label = ctk.CTkLabel(main_frame, text="No file selected", font=("Arial", 12))
    file_label.grid(row=1, column=1, columnspan=3, padx=10, pady=10, sticky="ew")

    # Sensors Frame
    sensors_frame = ctk.CTkFrame(main_frame, border_width=1, border_color="gray")
    sensors_frame.grid(row=2, column=0, pady=20, padx=20, sticky="nsew")
    # this must be updated later on to know which slider contains what.
    sliders, settings_sliders, pattern_params = create_sensors_frame(sensors_frame)

    # Events Frame
    events_frame = ctk.CTkFrame(main_frame, border_width=1, border_color="gray")
    events_frame.grid(row=2, column=1, columnspan=3, pady=20, padx=20, sticky="nsew")

    # Label for Events frame
    events_label = ctk.CTkLabel(events_frame, text="Events", font=("Arial", 14, "bold"))
    events_label.pack(pady=10)
    create_event_interface(events_frame)
    events = show_comments(events_frame)
    

    # Bottom Buttons
    button_export = ctk.CTkButton(main_frame, text="Export", command=lambda: exportToXlsx(df, file_name, sliders, events, settings_sliders, pattern_params), state='disabled')
    button_export.grid(row=3, column=0, columnspan=3, pady=10, sticky="ew")

    button_back = ctk.CTkButton(main_frame, text="Back", command=lambda: go_back_func(root, create_main_screen_func))
    button_back.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")

    # Configure grid weights for responsiveness
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_columnconfigure(2, weight=1)
    main_frame.grid_rowconfigure(2, weight=1)
    