import customtkinter as ctk
from utils import validateTime, convertTime, convertTimeToText, show_info_popup

commentsDict = {}

firstEventText = None

def placeComment(settings_frame):
    global commentsDict, hourText, minText, secText, commentText
    time = get_time_text()
    comment = commentText.get()
    if validateTime(time):
        commentsDict[convertTime(time)] = comment
    else:
        show_info_popup("Error", "You must enter the right format of time (HH:MM:SS)", settings_frame)
    # Only clear the event name field, keep time fields
    commentText.delete(0, ctk.END)

def create_event_interface(settings_frame):
    global hourText, minText, secText, commentText, firstEventText

    # Time and Comment Frame
    timecommentBundle = ctk.CTkFrame(settings_frame, fg_color="transparent")
    timecommentBundle.pack(padx=20, pady=20)

    # First Event Entry
    firstEventText = ctk.CTkEntry(timecommentBundle, width=300, placeholder_text="First-Event")
    firstEventText.pack(padx=2, pady=5)

    # Comment Entry (no label)
    commentText = ctk.CTkEntry(timecommentBundle, width=300, placeholder_text="Event")
    commentText.pack(padx=2, pady=5)

    # Frame to center time entry fields
    timeEntryFrame = ctk.CTkFrame(timecommentBundle, fg_color="transparent")
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

    # Create a frame to contain buttons and scrollable comments
    button_and_comments_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
    button_and_comments_frame.pack(fill="both", expand=True, padx=15, pady=15)

    # Place Event Button
    placeCommentButton = ctk.CTkButton(button_and_comments_frame, text="Place Event", 
                                     command=lambda: (placeComment(settings_frame), show_comments(settings_frame)))
    placeCommentButton.pack(pady=(10, 5), padx=10, fill="x")

    # Reset Events Button
    resetEventsButton = ctk.CTkButton(button_and_comments_frame, text="Reset Events", 
                                    command=lambda: reset_events(settings_frame))
    resetEventsButton.pack(pady=(0, 10), padx=10, fill="x")

    # Create scrollable frame for comments with proper sizing
    settings_frame.comments_frame = ctk.CTkScrollableFrame(button_and_comments_frame, 
                                                         height=150,  # Fixed height to prevent sizing issues
                                                         corner_radius=5)
    settings_frame.comments_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    return None

def get_time_text():
    hour = hourText.get().strip()
    minute = minText.get().strip()
    second = secText.get().strip()
    time_text = f"{hour}:{minute}:{second}"
    return time_text

def delete_comment(key, settings_frame):
    if key in commentsDict:
        del commentsDict[key]
    show_comments(settings_frame)

def edit_comment(key, settings_frame):
    if key in commentsDict:
        # Create edit popup
        edit_window = ctk.CTkToplevel()
        edit_window.title("Edit Event")
        edit_window.geometry("300x200")
        edit_window.transient(settings_frame.master)
        
        # Event name
        ctk.CTkLabel(edit_window, text="Event Name:").pack(pady=5)
        event_entry = ctk.CTkEntry(edit_window, width=250)
        event_entry.pack(pady=5)
        event_entry.insert(0, commentsDict[key])
        
        # Time
        ctk.CTkLabel(edit_window, text="Time (HH:MM:SS):").pack(pady=5)
        time_frame = ctk.CTkFrame(edit_window, fg_color="transparent")
        time_frame.pack(pady=5)
        
        current_time = convertTimeToText(key)
        time_parts = current_time.split(":")
        
        hour_entry = ctk.CTkEntry(time_frame, width=40)
        hour_entry.pack(side=ctk.LEFT, padx=2)
        hour_entry.insert(0, time_parts[0])
        
        ctk.CTkLabel(time_frame, text=":").pack(side=ctk.LEFT, padx=2)
        
        min_entry = ctk.CTkEntry(time_frame, width=40)
        min_entry.pack(side=ctk.LEFT, padx=2)
        min_entry.insert(0, time_parts[1])
        
        ctk.CTkLabel(time_frame, text=":").pack(side=ctk.LEFT, padx=2)
        
        sec_entry = ctk.CTkEntry(time_frame, width=40)
        sec_entry.pack(side=ctk.LEFT, padx=2)
        sec_entry.insert(0, time_parts[2])
        
        def save_edit():
            new_time = f"{hour_entry.get()}:{min_entry.get()}:{sec_entry.get()}"
            new_event = event_entry.get()
            
            if validateTime(new_time):
                # Remove old entry
                del commentsDict[key]
                # Add new entry
                commentsDict[convertTime(new_time)] = new_event
                edit_window.destroy()
                show_comments(settings_frame)
            else:
                show_info_popup("Error", "Invalid time format (HH:MM:SS)", edit_window)
        
        ctk.CTkButton(edit_window, text="Save", command=save_edit).pack(pady=10)

def reset_events(settings_frame):
    global commentsDict, firstEventText
    commentsDict.clear()
    
    # Clear input fields
    commentText.delete(0, ctk.END)
    hourText.delete(0, ctk.END)
    minText.delete(0, ctk.END)
    secText.delete(0, ctk.END)

    # Clear first event field
    if firstEventText:
        firstEventText.delete(0, ctk.END)
    
    show_comments(settings_frame)

def get_first_event_name():
    global firstEventText
    if firstEventText and firstEventText.get().strip():
        return firstEventText.get().strip()
    else:
        return "Post-Wake"

def show_comments(settings_frame):
    # Clear existing comments frame
    for widget in settings_frame.comments_frame.winfo_children():
        widget.destroy()

    # Show comments
    for key, value in commentsDict.items():
        comment_frame = ctk.CTkFrame(settings_frame.comments_frame)
        comment_frame.pack(pady=3, padx=3, fill='x')
        
        # Configure the grid for proper button positioning
        comment_frame.grid_columnconfigure(0, weight=1)  # Text expands
        comment_frame.grid_columnconfigure(1, weight=0)  # Buttons stay fixed

        # Create the full text
        full_text = f"Time: {convertTimeToText(key)} - Event: {value}"
        
        # Event text on the left with text wrapping and ellipsis
        timeAndCommentText = ctk.CTkLabel(comment_frame, 
                                        text=full_text,
                                        anchor="w",
                                        wraplength=260,  # Adjust this value as needed
                                        justify="left")
        timeAndCommentText.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="ew")

        # Button frame on the right with fixed position
        button_frame = ctk.CTkFrame(comment_frame, fg_color="transparent")
        button_frame.grid(row=0, column=1, padx=(5, 10), pady=5, sticky="e")

        # Only Edit and Delete buttons (Copy button removed)
        edit_button = ctk.CTkButton(button_frame, text="Edit", width=50,
                                  command=lambda k=key, sf=settings_frame: edit_comment(k, sf))
        edit_button.pack(side='left', padx=2)

        delete_button = ctk.CTkButton(button_frame, text="Delete", width=50,
                                    command=lambda k=key, sf=settings_frame: delete_comment(k, sf))
        delete_button.pack(side='left', padx=2)
    
    return commentsDict, firstEventText if 'firstEventText' in globals() else None