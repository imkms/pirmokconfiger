import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
from docx import Document
from PIL import Image, ImageTk

def extract_text(doc_path):
    """Extract all text from a Word document."""
    document = Document(doc_path)
    return "\n".join(para.text for para in document.paragraphs)

def find_available_options(text):
    """Find available device names in the text."""
    options = {"ASA", "RA", "Router1", "Router", "Router2", "RB", "S1", "SA", "SB", "S2"}
    found_options = {opt for opt in options if opt.lower() in text.lower()}
    return sorted(found_options)

def add_option():
    """Add the typed option to the dropdown menu."""
    new_option = option_entry.get().strip()
    if new_option and new_option.lower() not in [option.lower() for option in dropdown_menu['values']]:
        dropdown_menu['values'] = list(dropdown_menu['values']) + [new_option]
        dropdown_menu.set(new_option)
    else:
        # If the option is empty or already exists, display an error message
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, "Option already exists or is empty.")

def open_file():
    """Open a Word document and extract text."""
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        extracted_text = extract_text(file_path)
        
        # Static options predefined in the code
        available_options = ["ASA", "RA", "Router1", "Router", "Router2", "RB", "S1", "SA", "SB", "S2"]
        
        dropdown_menu['values'] = available_options  # Set available options statically
        dropdown_menu.set(available_options[0])  # Set the default selection
        
        global full_text  # Store the full extracted text for later use
        full_text = extracted_text

def confirm_choice():
    """Lock in the choice selected in the dropdown and filter lines."""
    selected_option = dropdown_menu.get()
    if selected_option != "No options found":
        filtered_lines = []
        for line in full_text.split('\n'):
            if line.lower().startswith(selected_option.lower()):
                filtered_lines.append(line)
        filtered_text = "\n".join(filtered_lines) if filtered_lines else "No matching lines found."
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, filtered_text)
    else:
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, "No valid option selected.")

def remove_text_before_hash():
    """Remove text before '#' for the selected option."""
    selected_option = dropdown_menu.get()
    if selected_option != "No options found":
        modified_lines = []
        for line in full_text.split('\n'):
            if line.lower().startswith(selected_option.lower()):
                # Remove everything before the '#' including the '#'
                index = line.find('#')
                if index != -1:
                    line = line[index + 1:].strip()  # Keep text after '#'
                modified_lines.append(line)
        
        modified_text = "\n".join(modified_lines) if modified_lines else "No matching lines found."
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, modified_text)
    else:
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, "No valid option selected.")

def toggle_magic():
    """Toggle the 'Magic' feature and hide or show the text."""
    global magic_on
    selected_option = dropdown_menu.get()

    # Only show filtered lines based on the selected option
    if selected_option != "No options found":
        # Store the original lines when Magic is first toggled on
        if not magic_on:
            # Store the unmodified filtered lines if it's the first toggle
            global unmodified_lines
            unmodified_lines = []
            for line in full_text.split('\n'):
                if line.lower().startswith(selected_option.lower()):
                    unmodified_lines.append(line)

        filtered_lines = []

        if magic_on:
            # If Magic is ON, modify the lines by removing text before the '#' if it exists
            for line in unmodified_lines:  # Use unmodified lines for Magic ON
                if '#' in line:
                    # Remove everything before and including the '#'
                    index = line.find('#')
                    line = line[index + 1:].strip()  # Keep text after '#'
                    filtered_lines.append(line)
                else:
                    # If no '#' is found, skip this line
                    continue
        else:
            # If Magic is OFF, use the unmodified lines as they were
            filtered_lines = unmodified_lines

        # Update the displayed text based on the "Magic" state
        filtered_text = "\n".join(filtered_lines) if filtered_lines else "No matching lines found."
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, filtered_text)

    # Change the button text based on the "Magic" state
    if magic_on:
        magic_btn.config(text="Magic: ON (spammeld picit mert buggos)")
    else:
        magic_btn.config(text="Magic: OFF (spammeld picit mert buggos(")

    # Toggle the magic state
    magic_on = not magic_on


def select_background_image():
    """Allow the user to select a background image and display it."""
    file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.jpg;*.png")])
    if file_path:
        image = Image.open(file_path)
        image = image.resize((root.winfo_width(), root.winfo_height()))  # Resize to fit the window
        bg_image = ImageTk.PhotoImage(image)
        
        # Create a Label widget for the background image
        background_label.config(image=bg_image)
        background_label.image = bg_image  # Keep a reference to avoid garbage collection

# GUI setup
root = tk.Tk()
root.title("Pirmok Gyors Configoloja (Alpha)")
root.geometry("800x600")  # Set window size

# Background Label (for image display)
background_label = tk.Label(root)
background_label.grid(row=0, column=0, columnspan=2, rowspan=2, sticky="nsew")

# Open Word Document button
btn_open = tk.Button(root, text="(1.)Valassz Pirmok Doksit                            (2.)->", command=open_file)
btn_open.grid(row=0, column=0, padx=3, pady=3, sticky="ew")

# Dropdown menu
dropdown_menu = ttk.Combobox(root, state="readonly")
dropdown_menu.grid(row=0, column=1, padx=3, pady=3, sticky="ew")

# Confirm Selection button
confirm_btn = tk.Button(root, text="(3.)Kivalasztottam a keszuleket amit keresek!", command=confirm_choice)
confirm_btn.grid(row=1, column=0, padx=3, pady=3, sticky="ew")

# Remove Text Before '#' button
magic_btn = tk.Button(root, text="Magic: OFF", command=toggle_magic)
magic_btn.grid(row=1, column=1, padx=3, pady=3, sticky="ew")


# Add new option entry and button
option_entry = tk.Entry(root)
option_entry.grid(row=2, column=0, padx=3, pady=3, sticky="ew")

add_option_btn = tk.Button(root, text="Sajat opcio hozzaadasa", command=add_option)
add_option_btn.grid(row=2, column=1, padx=3, pady=3, sticky="ew")

# Text display area
text_display = scrolledtext.ScrolledText(root, width=80, height=20)
text_display.grid(row=3, column=0, columnspan=2, padx=3, pady=3)

# Make all grid cells expand with window resizing
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_rowconfigure(3, weight=3)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Initialize magic state
magic_on = False

root.mainloop()
