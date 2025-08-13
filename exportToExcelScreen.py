import customtkinter as ctk
from exportToExcelScreen.export import exportToXlsx
from utils import clear_screen
from exportToExcelScreen.events import create_event_interface, show_comments
from exportToExcelScreen.sensors import create_sensors_frame
from exportToExcelScreen.importFile import select_input_file


def export_to_excel_screen(root, go_back_func, create_main_screen_func):
    clear_screen(root)

    # Main frame
    main_frame = ctk.CTkFrame(root)
    main_frame.pack(pady=20, padx=20, fill="both", expand=True)

    # Titel
    title_label = ctk.CTkLabel(main_frame, text="Data Analysis", font=("Arial", 20, "bold"))
    title_label.grid(row=0, column=0, columnspan=3, pady=10)

    # Huidige file-status
    df = None
    file_name = None

    # Helper: events UI (fallback als show_comments geen (dict, entry) teruggeeft)
    def build_events_ui(parent):
        events_dict = {}
        first_event_entry = ctk.CTkEntry(parent, placeholder_text="First event (bv. Post-Wake)")
        first_event_entry.pack(pady=(0, 6))
        comments_box = ctk.CTkTextbox(parent, width=360, height=100, state="disabled")
        comments_box.pack(pady=4)
        return events_dict, first_event_entry

    # Maak alvast een (nog lege) Export-knop aan zodat select_file hem kan enablen
    button_export = ctk.CTkButton(main_frame, text="Export", state='disabled')

    # File select
    file_label = ctk.CTkLabel(main_frame, text="No file selected", font=("Arial", 12))
    file_label.grid(row=1, column=1, columnspan=3, padx=10, pady=10, sticky="ew")

    def select_file_and_update_label():
        nonlocal df, file_name
        df, file_name = select_input_file(root, file_label, button_export)

    button_select_input = ctk.CTkButton(main_frame, text="Select Input File",
                                        command=select_file_and_update_label)
    button_select_input.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    # Sensors
    sensors_frame = ctk.CTkFrame(main_frame, border_width=1, border_color="gray")
    sensors_frame.grid(row=2, column=0, pady=20, padx=20, sticky="nsew")
    ctk.CTkLabel(sensors_frame, text="Sensors", font=("Arial", 14, "bold")).pack(pady=10)

    # sliders, settings_sliders, reset_sensors komen uit sensors.py
    sliders, settings_sliders, reset_sensors = create_sensors_frame(sensors_frame)

    # Detection settings (NIEUW): maak detect_frame AAN voordat je entries maakt
    detect_frame = ctk.CTkFrame(main_frame, border_width=1, border_color="gray")
    detect_frame.grid(row=2, column=2, pady=20, padx=20, sticky="nsew")
    ctk.CTkLabel(detect_frame, text="Detection settings", font=("Arial", 14, "bold")).pack(pady=(10, 6))

    entry_long_thr  = ctk.CTkEntry(detect_frame, placeholder_text="Long ≥ sensors", width=140)
    entry_long_thr.insert(0, "5"); entry_long_thr.pack(pady=4)

    entry_hapc_sens = ctk.CTkEntry(detect_frame, placeholder_text="HAPC/HARPC ≥ sensors > amp", width=140)
    entry_hapc_sens.insert(0, "3"); entry_hapc_sens.pack(pady=4)

    entry_hapc_amp  = ctk.CTkEntry(detect_frame, placeholder_text="Min amplitude (mmHg)", width=140)
    entry_hapc_amp.insert(0, "100"); entry_hapc_amp.pack(pady=4)

    # simpele numeric guard
    for e in (entry_long_thr, entry_hapc_sens, entry_hapc_amp):
        def _only_digits(event, widget=e):
            txt = widget.get()
            if txt and not txt.replace('.', '', 1).isdigit():
                widget.delete(0, "end")
                widget.insert(0, "".join(ch for ch in txt if ch.isdigit() or ch == '.'))
        e.bind("<KeyRelease>", _only_digits)

    settings = {
        "distance_between_sensors": settings_sliders[0],   # bestaande slider (al aanwezig)
        "long_threshold_sensors": entry_long_thr,
        "hapc_min_sensors": entry_hapc_sens,
        "hapc_min_amplitude": entry_hapc_amp,
    }

    # Events
    events_frame = ctk.CTkFrame(main_frame, border_width=1, border_color="gray")
    events_frame.grid(row=2, column=1, pady=20, padx=20, sticky="nsew")
    ctk.CTkLabel(events_frame, text="Events", font=("Arial", 14, "bold")).pack(pady=10)

    reset_events = create_event_interface(events_frame)

    # Gebruik show_comments als die (dict, entry) teruggeeft; anders fallback
    try:
        maybe = show_comments(events_frame)
        if isinstance(maybe, tuple) and len(maybe) == 2:
            events, first_event_entry = maybe
        else:
            events, first_event_entry = build_events_ui(events_frame)
    except Exception:
        events, first_event_entry = build_events_ui(events_frame)

    # Export-knop nu pas configureren en plaatsen
    button_export.configure(
        command=lambda: exportToXlsx(df, file_name, sliders, events, settings_sliders, first_event_entry, settings),
        state='normal'
    )
    button_export.grid(row=3, column=0, columnspan=3, pady=10, sticky="ew")

    # Reset/back
    def reset_sensors_all():
        reset_sensors()
        entry_long_thr.delete(0, "end"); entry_long_thr.insert(0, "5")
        entry_hapc_sens.delete(0, "end"); entry_hapc_sens.insert(0, "3")
        entry_hapc_amp.delete(0, "end");  entry_hapc_amp.insert(0, "100")

    def back_and_reset():
        reset_sensors_all()
        reset_events()
        go_back_func(root, create_main_screen_func)

    button_back = ctk.CTkButton(main_frame, text="Back", command=back_and_reset)
    button_back.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")

    # Grid weights
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_columnconfigure(2, weight=1)
    main_frame.grid_rowconfigure(2, weight=1)
