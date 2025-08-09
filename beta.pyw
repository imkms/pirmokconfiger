import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, font
from docx import Document
from PIL import Image, ImageTk, ImageOps # Import ImageOps for potential future use or better image handling

# --- Global Variables ---
full_text = ""
magic_on = False
# These variables will be assigned to Tkinter widgets later in the GUI setup
root = None
dropdown_menu = None
text_display = None
option_entry = None
magic_btn = None
background_label = None

# --- Helper Functions ---
def extract_text(doc_path):
    """Extract all text from a Word document."""
    document = Document(doc_path)
    return "\n".join(para.text for para in document.paragraphs)

def find_available_options(text):
    """Find available device names in the text."""
    options = {"ASA", "RA", "Router1", "Router", "Router2", "RB", "S1", "SA", "SB", "S2"}
    found_options = {opt for opt in options if opt.lower() in text.lower()}
    return sorted(list(found_options))

# --- Event Handlers ---
def add_option():
    """Add the typed option to the dropdown menu."""
    new_option = option_entry.get().strip()
    if new_option and new_option.lower() not in [str(v).lower() for v in dropdown_menu['values']]:
        dropdown_menu['values'] = list(dropdown_menu['values']) + [new_option]
        dropdown_menu.set(new_option)
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, f"'{new_option}' added to options.")
    else:
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, "Option already exists or is empty.")

def open_file():
    """Open a Word document, extract text, and find device options."""
    global full_text
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        full_text = extract_text(file_path)
        
        available_options = find_available_options(full_text)
        
        if available_options:
            dropdown_menu['values'] = available_options
            dropdown_menu.set(available_options[0])
        else:
            dropdown_menu['values'] = ["No options found"]
            dropdown_menu.set("No options found")
        
        confirm_choice() # Automatically show config for the first found option

def confirm_choice():
    """Lock in the choice selected in the dropdown and filter lines."""
    selected_option = dropdown_menu.get()
    if selected_option != "No options found" and full_text:
        # Ensure magic is OFF when confirming a new choice
        global magic_on
        magic_on = False
        magic_btn.config(text="Magic: OFF")

        filtered_lines = [line for line in full_text.split('\n') if line.lower().startswith(selected_option.lower())]
        filtered_text = "\n".join(filtered_lines) if filtered_lines else "No matching lines found."
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, filtered_text)
    else:
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, "No valid option selected or no file opened.")

def toggle_magic():
    """Toggle the 'Magic' feature to show/hide configuration commands."""
    global magic_on

    selected_option = dropdown_menu.get()
    if selected_option == "No options found" or not full_text:
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, "No valid option selected or no file opened.")
        return

    # Get all lines from the full text that start with the selected option
    base_lines = [line for line in full_text.split('\n') if line.lower().startswith(selected_option.lower())]
    
    output_lines = []
    if not magic_on: # If magic is currently OFF, we are turning it ON
        for line in base_lines:
            if '#' in line:
                index = line.find('#')
                output_lines.append(line[index + 1:].strip())
            # Lines without '#' are skipped when magic is ON
    else: # If magic is currently ON, we are turning it OFF
        output_lines = base_lines # Show original lines for the selected option

    filtered_text = "\n".join(output_lines) if output_lines else "No matching lines found."
    text_display.delete(1.0, tk.END)
    text_display.insert(tk.END, filtered_text)

    # Toggle the state AFTER applying the logic for the *current* state
    magic_on = not magic_on
    magic_btn.config(text=f"Magic: {'ON' if magic_on else 'OFF'}")

def select_background_image():
    """Allow the user to select a background image and display it."""
    file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.jpg;*.jpeg;*.png")])
    if file_path:
        try:
            image = Image.open(file_path)
            # Resize image to fit the current window dimensions
            # Use Image.Resampling.LANCZOS for high-quality downsampling
            image = image.resize((root.winfo_width(), root.winfo_height()), Image.Resampling.LANCZOS)
            bg_image = ImageTk.PhotoImage(image)
            
            background_label.config(image=bg_image)
            background_label.image = bg_image # Keep a reference to avoid garbage collection
            background_label.lower() # Ensure background label is behind other widgets
        except Exception as e:
            text_display.delete(1.0, tk.END)
            text_display.insert(tk.END, f"Error loading image: {e}")

# --- GUI Setup ---
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Pirmok's Quick Configurator")
    root.geometry("850x650")
    root.resizable(False, False) # User requested fixed size
    root.configure(background="#2E2E2E") # Fallback background color

    # --- Style Configuration ---
    BG_COLOR = "#2E2E2E" # Dark background
    FG_COLOR = "#EAEAEA" # Light foreground
    BTN_BG = "#4A4A4A" # Darker button background
    BTN_FG = "#FFFFFF" # White button text
    ACCENT_COLOR = "#007ACC" # Blue accent for hover/selection
    FONT_FAMILY = "Segoe UI" # Common Windows font, good readability
    FONT_NORMAL = (FONT_FAMILY, 11)
    FONT_BOLD = (FONT_FAMILY, 11, "bold")

    style = ttk.Style(root)
    style.theme_use('clam') # 'clam' theme is highly customizable

    # Configure general styles
    style.configure('TFrame', background=BG_COLOR)
    style.configure('TButton', 
                    background=BTN_BG, 
                    foreground=BTN_FG, 
                    font=FONT_BOLD, 
                    padding=10, 
                    borderwidth=0, 
                    relief='flat')
    style.map('TButton', 
              background=[('active', ACCENT_COLOR)], # Hover effect
              foreground=[('active', BTN_FG)]) # Keep text color white on hover

    style.configure('TCombobox', 
                    selectbackground=ACCENT_COLOR, # Background of selected item in dropdown list
                    fieldbackground=BTN_BG, # Background of the combobox entry field
                    background=BTN_BG, # Background of the dropdown button
                    foreground=FG_COLOR, 
                    arrowcolor=FG_COLOR, 
                    font=FONT_NORMAL,
                    borderwidth=1,
                    relief='flat')
    # Configure the dropdown listbox itself
    root.option_add('*TCombobox*Listbox.background', BTN_BG)
    root.option_add('*TCombobox*Listbox.foreground', FG_COLOR)
    root.option_add('*TCombobox*Listbox.selectBackground', ACCENT_COLOR)
    root.option_add('*TCombobox*Listbox.font', FONT_NORMAL)

    style.configure('TEntry', 
                    fieldbackground=BTN_BG, 
                    foreground=FG_COLOR, 
                    insertcolor=FG_COLOR, 
                    font=FONT_NORMAL, 
                    borderwidth=1, 
                    relief='flat')
    style.configure('TLabel', background=BG_COLOR, foreground=FG_COLOR, font=FONT_NORMAL)

    # --- Background and Main Frame ---
    # Place background_label first to be behind other widgets
    background_label = tk.Label(root, bg=BG_COLOR)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)

    # Content frame to hold all other widgets, placed on top of background_label
    content_frame = ttk.Frame(root, style='TFrame', padding=20)
    content_frame.pack(fill="both", expand=True) # Pack to fill the root window

    # Configure grid for content_frame
    content_frame.grid_columnconfigure((0, 1), weight=1)
    content_frame.grid_rowconfigure(3, weight=1) # Text display row expands vertically

    # --- Widgets ---
    btn_open = ttk.Button(content_frame, text="(1) Select Document", command=open_file)
    btn_open.grid(row=0, column=0, padx=5, pady=10, sticky="ew")

    dropdown_menu = ttk.Combobox(content_frame, state="readonly", font=FONT_NORMAL)
    dropdown_menu.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
    dropdown_menu.set("Open a document to see devices")

    confirm_btn = ttk.Button(content_frame, text="(2) Show Config for Device", command=confirm_choice)
    confirm_btn.grid(row=1, column=0, padx=5, pady=10, sticky="ew")

    magic_btn = ttk.Button(content_frame, text="Magic: OFF", command=toggle_magic)
    magic_btn.grid(row=1, column=1, padx=5, pady=10, sticky="ew")

    option_entry = ttk.Entry(content_frame, font=FONT_NORMAL)
    option_entry.grid(row=2, column=0, padx=5, pady=10, sticky="ew")

    add_option_btn = ttk.Button(content_frame, text="Add Custom Device", command=add_option)
    add_option_btn.grid(row=2, column=1, padx=5, pady=10, sticky="ew")

    text_display = scrolledtext.ScrolledText(content_frame, 
                                             width=80, 
                                             height=20, 
                                             bg=BTN_BG, # Use button background for text area
                                             fg=FG_COLOR, 
                                             font=("Consolas", 10), # Monospace font for code display
                                             relief='flat', 
                                             borderwidth=0, 
                                             insertbackground=FG_COLOR, 
                                             padx=10, 
                                             pady=10)
    text_display.grid(row=3, column=0, columnspan=2, padx=5, pady=10, sticky="nsew")

    bg_btn = ttk.Button(content_frame, text="Change Background", command=select_background_image)
    bg_btn.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

    root.mainloop()
