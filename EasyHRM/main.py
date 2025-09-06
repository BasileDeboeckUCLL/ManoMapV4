import customtkinter as ctk
from exportToExcelScreen.exportToExcelScreen import export_to_excel_screen
from patternDetectionScreen.patternDetectionScreen import open_screen_for_pattern_detection
from utils import go_back, toggle_mode
import sys
import os
from PIL import Image

import sys
import os
import customtkinter as ctk
from exportToExcelScreen.exportToExcelScreen import export_to_excel_screen
from patternDetectionScreen.patternDetectionScreen import open_screen_for_pattern_detection
from utils import go_back, toggle_mode
from PIL import Image

import utils

import utils

def create_main_window():
    app = ctk.CTk()
    app.title("EasyHRM")

    def resource_path(relative_path):
        """ Get the absolute path to the resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath("./EasyHRM")
        return os.path.join(base_path, relative_path)

    icon_path = resource_path("EasyHRM_icon.ico")
    app.iconbitmap(icon_path)

    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()

    window_width = int(screen_width * 0.8)
    window_height = int(screen_height * 0.85)

    position_top = int(screen_height / 2 - window_height / 2) - 30
    position_right = int(screen_width / 2 - window_width / 2)

    app.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

    def on_configure(event):
        if app.winfo_exists():
            utils.window_width = event.width
            utils.window_height = event.height
            utils.window_x = app.winfo_x()
            utils.window_y = app.winfo_y()

    app.bind("<Configure>", on_configure)

    return app

def build_main_screen(app):
    utils.clear_screen(app)

    def resource_path(relative_path):
        """ Get the absolute path to the resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath("./EasyHRM")
        return os.path.join(base_path, relative_path)

    main_frame = ctk.CTkFrame(app, corner_radius=0)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    title_logo_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    title_logo_frame.pack(pady=20)

    title_label = ctk.CTkLabel(title_logo_frame, text="EasyHRM", font=("Arial", 30, "bold"))
    title_label.pack(side="left", padx=10)

    logo_path = resource_path("EasyHRM_icon.ico")
    logo_image = Image.open(logo_path)
    logo = ctk.CTkImage(logo_image, size=(75, 75))
    logo_label = ctk.CTkLabel(title_logo_frame, image=logo, text="")
    logo_label.pack(side="left", padx=10)

    description_text = (
        "Optimise your colon examination with our application! "
        "Automate the time-consuming process of identifying colon patterns. "
        "Save valuable time for examination and analysis, and improve the accuracy of your data."
    )
    description_label = ctk.CTkLabel(main_frame, text=description_text, font=("Arial", 14), wraplength=600, justify="center")
    description_label.pack(pady=25)

    instruction_label = ctk.CTkLabel(main_frame, text="What would you like to do?", font=("Arial", 18, "bold"))
    instruction_label.pack(pady=10)

    button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    button_frame.pack(pady=20)

    button_a = ctk.CTkButton(button_frame, text="Pattern Detection", command=lambda: open_screen_for_pattern_detection(app, go_back, lambda: build_main_screen(app)), width=240, height=50, font=("Arial", 14, "bold"))
    button_a.pack(pady=10)

    button_b = ctk.CTkButton(button_frame, text="Data Analysis", command=lambda: export_to_excel_screen(app, go_back, lambda: build_main_screen(app)), width=240, height=50, font=("Arial", 14, "bold"))
    button_b.pack(pady=10)

    mode_toggle_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    mode_toggle_frame.pack(pady=20)

    initial_mode = ctk.get_appearance_mode().strip().lower() == "dark"
    mode_toggle = ctk.CTkSwitch(mode_toggle_frame, text="Light / Dark mode", command=toggle_mode, onvalue=1, offvalue=0, font=("Arial", 14, "bold"))
    mode_toggle.pack()
    mode_toggle.select() if initial_mode else mode_toggle.deselect()

# Run the main screen
if __name__ == "__main__":
    app = create_main_window()
    build_main_screen(app)
    app.mainloop()
