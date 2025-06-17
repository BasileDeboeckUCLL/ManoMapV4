import customtkinter as ctk
from utils import validateTime, convertTime, convertTimeToText, show_info_popup

commentsDict = {}


def placeComment(settings_frame):
    # Add hourText, minText, secText, commentText as globals
    global commentsDict, hourText, minText, secText, commentText
    time = get_time_text()
    comment = commentText.get()
    if validateTime(time):
        commentsDict[convertTime(time)] = comment
        # Clear the event name and timestamp fields after adding
        commentText.delete(0, ctk.END)
        hourText.delete(0, ctk.END)
        minText.delete(0, ctk.END)
        secText.delete(0, ctk.END)
        # Remove focus from all input fields
        commentText.master.focus_set()
        # show_info_popup("Event", f"Event placed at {time}", settings_frame)
        # Pop up comment eruit gehaald
    else:
        show_info_popup(
            "Error", "You must enter the right format of time (HH:MM:SS)", settings_frame)


def create_event_interface(settings_frame):
    global hourText, minText, secText, commentText  # Declare globals

    def reset_events():
        global commentsDict
        commentsDict.clear()
        show_comments(settings_frame)

    # Reset Events Button
    reset_button = ctk.CTkButton(
        settings_frame, text="Reset Events", command=reset_events)
    reset_button.pack(pady=5)

    # Time and Comment Frame
    timecommentBundle = ctk.CTkFrame(settings_frame)
    timecommentBundle.pack(padx=20, pady=20)

    # Event Label
    commentText = ctk.StringVar()
    event_label = ctk.CTkLabel(timecommentBundle, textvariable=commentText)
    commentText.set("Event")
    event_label.pack()

    # Comment Entry
    commentText = ctk.CTkEntry(
        timecommentBundle, width=300, placeholder_text="Event")
    commentText.pack(padx=2, pady=5)

    # Time Label
    timeText = ctk.StringVar()
    time_label = ctk.CTkLabel(timecommentBundle, textvariable=timeText)
    timeText.set("Time")
    time_label.pack()

    # Frame to center time entry fields
    timeEntryFrame = ctk.CTkFrame(timecommentBundle)
    timeEntryFrame.pack(pady=5)

    # Hour Entry
    hourText = ctk.CTkEntry(timeEntryFrame, width=40, placeholder_text="HH")
    hourText.pack(side=ctk.LEFT, padx=2)

    # Separator
    colon1 = ctk.CTkLabel(timeEntryFrame, text=":")
    colon1.pack(side=ctk.LEFT, padx=2)

    # Minute Entry
    minText = ctk.CTkEntry(timeEntryFrame, width=40, placeholder_text="MM")
    minText.pack(side=ctk.LEFT, padx=2)

    # Separator
    colon2 = ctk.CTkLabel(timeEntryFrame, text=":")
    colon2.pack(side=ctk.LEFT, padx=2)

    # Second Entry
    secText = ctk.CTkEntry(timeEntryFrame, width=40, placeholder_text="SS")
    secText.pack(side=ctk.LEFT, padx=2)

    # Place Event Button with standard width
    placeCommentButton = ctk.CTkButton(settings_frame, text="Place Event", width=200,
                                       command=lambda: (placeComment(settings_frame), show_comments(settings_frame)))
    placeCommentButton.pack(pady=5)

    # Reset Events Button moved under Place Event
    reset_button.pack_forget()  # Remove previous packing
    reset_button.pack(pady=5)

    return reset_events


def get_time_text():
    # Get the content of each Entry widget and strip any extra whitespace
    hour = hourText.get().strip()
    minute = minText.get().strip()
    second = secText.get().strip()

    # Combine the time components
    time_text = f"{hour}:{minute}:{second}"
    return time_text


def delete_comment(key, label, settings_frame):
    del commentsDict[key]
    label.destroy()
    # Refresh comments frame
    show_comments(settings_frame)


def copy_comment(key, value, settings_frame):
    # Find a new unique key by incrementing seconds until unused
    new_key = key
    while new_key in commentsDict:
        new_key += 1  # increment by 1 second
    commentsDict[new_key] = value
    show_comments(settings_frame)


def edit_comment(key, value, settings_frame):
    # Create edit dialog
    edit_window = ctk.CTkToplevel()
    edit_window.title("Edit Event")
    edit_window.geometry("300x150")

    # Center the window relative to the main window
    main_window = settings_frame.winfo_toplevel()
    main_x = main_window.winfo_rootx()
    main_y = main_window.winfo_rooty()
    main_width = main_window.winfo_width()
    main_height = main_window.winfo_height()
    window_width = 300
    window_height = 150
    x = main_x + (main_width - window_width) // 2
    y = main_y + (main_height - window_height) // 2
    edit_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Make the window modal and on top
    edit_window.transient(main_window)
    edit_window.grab_set()
    edit_window.focus_force()

    # Time entry
    time_frame = ctk.CTkFrame(edit_window)
    time_frame.pack(pady=5)

    time_parts = convertTimeToText(key).split(':')

    hour_entry = ctk.CTkEntry(time_frame, width=40, placeholder_text="HH")
    hour_entry.insert(0, time_parts[0])
    hour_entry.pack(side='left', padx=2)

    ctk.CTkLabel(time_frame, text=":").pack(side='left')

    min_entry = ctk.CTkEntry(time_frame, width=40, placeholder_text="MM")
    min_entry.insert(0, time_parts[1])
    min_entry.pack(side='left', padx=2)

    ctk.CTkLabel(time_frame, text=":").pack(side='left')

    sec_entry = ctk.CTkEntry(time_frame, width=40, placeholder_text="SS")
    sec_entry.insert(0, time_parts[2])
    sec_entry.pack(side='left', padx=2)

    # Event name entry
    name_entry = ctk.CTkEntry(edit_window, width=200)
    name_entry.insert(0, value)
    name_entry.pack(pady=10)

    def save_changes():
        new_time = f"{hour_entry.get()}:{min_entry.get()}:{sec_entry.get()}"
        if validateTime(new_time):
            del commentsDict[key]
            commentsDict[convertTime(new_time)] = name_entry.get()
            show_comments(settings_frame)
            edit_window.destroy()
        else:
            show_info_popup(
                "Error", "Invalid time format (HH:MM:SS)", edit_window)

    save_button = ctk.CTkButton(edit_window, text="Save", command=save_changes)
    save_button.pack(pady=10)


def show_comments(settings_frame):
    # Create comments frame if it doesn't exist and there are comments
    if not hasattr(settings_frame, 'comments_frame') and commentsDict:
        settings_frame.comments_frame = ctk.CTkFrame(settings_frame)
        settings_frame.comments_frame.pack(padx=1, pady=1)
    elif hasattr(settings_frame, 'comments_frame'):
        # Clear existing comments frame
        for widget in settings_frame.comments_frame.winfo_children():
            widget.destroy()
        # Remove frame if no comments
        if not commentsDict:
            settings_frame.comments_frame.destroy()
            delattr(settings_frame, 'comments_frame')
            return commentsDict

    # Show comments
    for key, value in sorted(commentsDict.items()):
        comment_frame = ctk.CTkFrame(settings_frame.comments_frame)
        comment_frame.pack(pady=5, padx=10, fill='x')

        # Format text: Event name first, then time, bold text, more padding, slightly larger font
        text_label = ctk.CTkLabel(
            comment_frame, text=f"Event: {value} - Time: {convertTimeToText(key)}", font=("Arial", 12, "bold"))
        text_label.pack(side='left', padx=(10, 15), pady=5)

        # Buttons frame for better organization
        buttons_frame = ctk.CTkFrame(comment_frame)
        buttons_frame.pack(side='right', padx=5)

        copy_button = ctk.CTkButton(buttons_frame, text="Duplicate", width=60,
                                    command=lambda k=key, v=value, sf=settings_frame: copy_comment(k, v, sf))
        copy_button.pack(side='left', padx=5)

        edit_button = ctk.CTkButton(buttons_frame, text="Edit", width=60,
                                    command=lambda k=key, v=value, sf=settings_frame: edit_comment(k, v, sf))
        edit_button.pack(side='left', padx=5)

        delete_button = ctk.CTkButton(buttons_frame, text="Delete", width=60,
                                      command=lambda k=key, lbl=text_label, sf=settings_frame: delete_comment(k, lbl, sf))
        delete_button.pack(side='left', padx=5)
    return commentsDict
