import customtkinter as ctk
from utils import validateTime, convertTime, convertTimeToText, show_info_popup

commentsDict = {}

def placeComment(settings_frame):
    global commentsDict, hourText, minText, secText, commentText  # Add hourText, minText, secText, commentText as globals
    time = get_time_text()
    comment = commentText.get()
    if validateTime(time):
        commentsDict[convertTime(time)] = comment
        # show_info_popup("Event", f"Event placed at {time}", settings_frame)
        # Pop up comment eruit gehaald
    else:
        show_info_popup("Error", "You must enter the right format of time (HH:MM:SS)", settings_frame)
    commentText.delete(0, ctk.END)

def create_event_interface(settings_frame):
    global hourText, minText, secText, commentText  # Declare globals

    # Label for first comment, which has no specified time. Start time is just the start
    first_event_label = ctk.CTkLabel(settings_frame, text="First Event: ")
    first_event_label.pack(padx=1, pady=1)
    
    # Frame for the name of the first event
    first_event_text = ctk.CTkEntry(settings_frame)
    first_event_text.pack(padx=1, pady=1)

    # Time and Comment Frame
    timecommentBundle = ctk.CTkFrame(settings_frame)
    timecommentBundle.pack(padx=20, pady=20)

    # Create a frame to contain comments
    settings_frame.comments_frame = ctk.CTkFrame(settings_frame)
    settings_frame.comments_frame.pack(padx=1, pady=1)
    
    # Event Label
    commentText = ctk.StringVar()
    event_label = ctk.CTkLabel(timecommentBundle, textvariable=commentText)
    commentText.set("Event")
    event_label.pack()

    # Comment Entry
    commentText = ctk.CTkEntry(timecommentBundle, width=300, placeholder_text="Event")
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

    # Place Event Button
    placeCommentButton = ctk.CTkButton(settings_frame, text="Place Event", command=lambda: (placeComment(settings_frame), show_comments(settings_frame)))
    placeCommentButton.pack(pady=10, padx=10)
    return first_event_text

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

def show_comments(settings_frame):
    # Clear existing comments frame
    for widget in settings_frame.comments_frame.winfo_children():
        widget.destroy()

    # Show comments
    for key, value in commentsDict.items():
        comment_frame = ctk.CTkFrame(settings_frame.comments_frame)
        comment_frame.pack(pady=10, padx=10, fill='x')

        timeAndCommentText = ctk.CTkLabel(comment_frame, text=f"Time: {convertTimeToText(key)} - Event: {value}")
        timeAndCommentText.pack(side='left', padx=(10, 5), pady=5)

        delete_button = ctk.CTkButton(comment_frame, text="Delete", command=lambda k=key, lbl=timeAndCommentText, sf=settings_frame: delete_comment(k, lbl, sf))
        delete_button.pack(side='right', padx=(5, 10), pady=5)
    return commentsDict
