import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import uuid # Original for MAC address, kept as fallback
import hashlib # For hashing the machine ID to create an activation key
import json # For saving and loading activation status
import platform # To check OS for WMI
try:
    import wmi # For detailed Windows machine ID
except ImportError:
    wmi = None # Set to None if WMI is not available

# --- IMPORTANT: Secret Phrase for Activation Key Generation ---
# This MUST be IDENTICAL to the one in your Easy File Renamer App's key generator.
AUTHORIZED_SECRET_PHRASE = "MyCustomSecretPhraseForRenamerApp2025!"

# --- Global Constant for Activation File ---
# This file will store whether the application has been activated.
ACTIVATION_FILE = "activation_status.json"

class FolderCreatorApp:
    def __init__(self, master):
        self.master = master
        master.title("Indexing PRO")
        master.resizable(False, False) # Prevents user from resizing the window

        # --- Styling Configuration ---
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0", font=("Inter", 10))

        # Enhanced style for the main "Create Folders" button (black background, RED text)
        style.configure("TButton", font=("Inter", 14, "bold"), padding=15, borderwidth=3, relief="raised",
                        background="black", foreground="red")
        style.map("TButton",
                  background=[('active', '#333333'), ('!disabled', 'black')],
                  foreground=[('active', 'red'), ('!disabled', 'red')])

        # Specific style for the "Browse for Location" and "Copy Machine ID" buttons (red text)
        style.configure("Browse.TButton", foreground="red")
        style.map("Browse.TButton",
                  background=[('active', '#333333'), ('!disabled', 'black')],
                  foreground=[('active', 'red'), ('!disabled', 'red')])

        # Explicitly set TEntry foreground to black for visibility
        style.configure("TEntry", font=("Inter", 10), padding=5, foreground="black")
        style.configure("TProgressbar", thickness=10)
        style.configure("TCheckbutton", background="#f0f0f0", font=("Inter", 10))
        style.configure("TCombobox", font=("Inter", 10)) # Style for combobox

        # --- Check Activation Status on Startup ---
        # Load the activation status from the file.
        self.activated = self._load_activation_status()

        # Based on activation status, show either the activation screen or the main app UI
        if not self.activated:
            self._setup_activation_screen()
        else:
            self._setup_main_app_ui()

    def _get_machine_id(self):
        """
        Generates a unique machine ID (System Code) for the machine.
        Uses WMI for Windows for a more robust fingerprint.
        Falls back to MAC address if WMI is not available or on non-Windows OS.
        """
        if platform.system() == "Windows" and wmi:
            try:
                c = wmi.WMI()

                cpu_info = c.Win32_Processor()[0]
                cpu_id = cpu_info.ProcessorId if hasattr(cpu_info, 'ProcessorId') else ""

                board_info = c.Win32_BaseBoard()[0]
                board_serial = board_info.SerialNumber if hasattr(board_info, 'SerialNumber') else ""

                disk_serial = ""
                disk_info = c.Win32_DiskDrive()
                for disk in disk_info:
                    if not disk.MediaType == "Removable Media":
                        disk_serial = disk.SerialNumber
                        break

                mac_address = ""
                for nic in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                    if nic.MACAddress:
                        mac_address = nic.MACAddress
                        break

                raw_fingerprint_string = f"{cpu_id}-{board_serial}-{disk_serial}-{mac_address}".strip()
                # Hash the raw fingerprint string to create the system code
                return hashlib.sha256(raw_fingerprint_string.encode('utf-8')).hexdigest()

            except Exception as e:
                messagebox.showwarning("WMI Error", f"Could not retrieve full machine ID using WMI: {e}\nFalling back to MAC address. Ensure 'wmi' is installed (pip install wmi) and you have permissions.")
                # Fallback to MAC address if WMI fails
                return hex(uuid.getnode())[2:].upper()
        else:
            if platform.system() == "Windows" and not wmi:
                messagebox.showwarning("WMI Module Missing", "The 'wmi' module is not installed. For a more robust machine ID, please install it using: pip install wmi\nFalling back to MAC address.")
            elif platform.system() != "Windows":
                messagebox.showinfo("Platform Info", "Detailed machine ID generation (WMI) is primarily supported on Windows. Using MAC address as machine ID.")

            # Fallback to MAC address for non-Windows or if wmi is not installed
            return hex(uuid.getnode())[2:].upper()

    def _generate_expected_key(self, machine_id):
        """
        Generates the expected activation key based on the machine ID and the
        predefined secret phrase. This must match the key generation logic
        in your separate key generator utility.
        """
        if not machine_id:
            return None
        combined_string = f"{machine_id}-{AUTHORIZED_SECRET_PHRASE}"
        return hashlib.sha256(combined_string.encode('utf-8')).hexdigest()

    def _load_activation_status(self):
        """
        Loads the activation status from the `ACTIVATION_FILE`.
        Returns True if activated, False otherwise (including file not found or corrupted).
        """
        if os.path.exists(ACTIVATION_FILE):
            try:
                with open(ACTIVATION_FILE, 'r') as f:
                    data = json.load(f)
                    # Return the 'activated' value, default to False if key is missing
                    return data.get('activated', False)
            except (json.JSONDecodeError, KeyError):
                # Handle cases where the JSON file is invalid or missing the key
                print(f"Warning: Could not read activation status from {ACTIVATION_FILE}. Assuming not activated.")
                return False
        return False

    def _save_activation_status(self, activated_status):
        """
        Saves the current activation status to the `ACTIVATION_FILE`.
        """
        try:
            with open(ACTIVATION_FILE, 'w') as f:
                json.dump({'activated': activated_status}, f)
            self.activated = activated_status # Update internal state
        except IOError as e:
            print(f"Error saving activation status to {ACTIVATION_FILE}: {e}")
            messagebox.showerror("Save Error", f"Could not save activation status: {e}")

    def _setup_activation_screen(self):
        """
        Sets up the graphical user interface for the application activation process.
        This screen prompts the user to enter an activation key.
        """
        # Create a frame to hold all activation widgets
        self.activation_frame = ttk.Frame(self.master, padding="20 20 20 20")
        self.activation_frame.pack(expand=True, fill="both")

        # Title and instructions
        ttk.Label(self.activation_frame, text="Indexing PRO Activation", font=("Inter", 16, "bold")).pack(pady=10)
        ttk.Label(self.activation_frame, text="Please activate your software to continue using Indexing PRO.").pack(pady=5)

        # Display the unique Machine ID
        self.machine_id = self._get_machine_id()
        ttk.Label(self.activation_frame, text="Your Machine ID (System Code):").pack(pady=(10, 0))
        ttk.Label(self.activation_frame, text=self.machine_id, font=("Inter", 12, "bold"), foreground="blue", wraplength=400).pack(pady=(0, 10))

        # Button to copy the Machine ID to clipboard
        ttk.Button(self.activation_frame, text="Copy Machine ID", command=self._copy_machine_id, style="Browse.TButton").pack(pady=5)
        ttk.Label(self.activation_frame, text="Provide this ID to the software vendor to obtain your activation key.", wraplength=400).pack(pady=5)

        # Input field for the Activation Key
        ttk.Label(self.activation_frame, text="Enter Activation Key:").pack(pady=(10, 0))
        self.activation_key_var = tk.StringVar()
        self.activation_key_entry = ttk.Entry(self.activation_frame, textvariable=self.activation_key_var, width=50)
        self.activation_key_entry.pack(pady=(0, 10))
        # Bind the Enter key to trigger activation
        self.activation_key_entry.bind("<Return>", self._activate_app)
        self.activation_key_entry.focus_set() # Set focus to the entry field

        # Activate Button
        ttk.Button(self.activation_frame, text="Activate", command=self._activate_app).pack(pady=10)

        # Status label for activation messages (e.g., success, error)
        self.activation_status_label = ttk.Label(self.activation_frame, text="", foreground="red", font=("Inter", 10))
        self.activation_status_label.pack(pady=5)

    def _copy_machine_id(self):
        """
        Copies the generated machine ID to the user's clipboard.
        """
        try:
            self.master.clipboard_clear() # Clear existing clipboard content
            self.master.clipboard_append(self.machine_id) # Add machine ID to clipboard
            self.master.update() # Update the clipboard immediately
            self.activation_status_label.config(text="Machine ID copied to clipboard!", foreground="green")
        except Exception as e:
            self.activation_status_label.config(text=f"Error copying ID: {e}", foreground="red")
            messagebox.showerror("Clipboard Error", f"Failed to copy Machine ID to clipboard: {e}")

    def _activate_app(self, event=None):
        """
        Validates the entered activation key against the expected key generated
        from the current machine ID. If valid, activates the application.
        """
        entered_key = self.activation_key_var.get().strip()
        expected_key = self._generate_expected_key(self.machine_id)

        if entered_key == expected_key:
            self._save_activation_status(True) # Save activation status as True
            self.activation_status_label.config(text="Activation successful! Starting application...", foreground="green")
            messagebox.showinfo("Activation Success", "Indexing PRO has been successfully activated! You can now use the application.")
            self._show_main_app() # Transition to the main application UI
        else:
            self.activation_status_label.config(text="Invalid Activation Key. Please try again.", foreground="red")
            messagebox.showerror("Activation Failed", "The entered activation key is incorrect. Please double-check and try again.")

    def _show_main_app(self):
        """
        Destroys the activation screen and sets up the main application UI.
        This method is called after successful activation.
        """
        # Destroy all widgets in the activation frame
        for widget in self.activation_frame.winfo_children():
            widget.destroy()
        self.activation_frame.destroy() # Destroy the frame itself
        self._setup_main_app_ui() # Initialize the main application UI

    def _setup_main_app_ui(self):
        """
        Sets up all the UI components for the main folder creation application.
        This is the existing code from your original application, now integrated
        to appear after successful activation.
        """
        # Variables to store user inputs
        self.output_location = tk.StringVar()
        self.main_folder_code = tk.StringVar()
        self.num_sub_folders = tk.StringVar()

        # List to hold dictionaries for each book's input fields
        self.book_inputs = []

        self.is_fullscreen = tk.BooleanVar(value=True) # State for fullscreen toggle
        self.auto_pick_code = tk.BooleanVar(value=False) # Variable for auto-pick option

        # Flag to indicate if WF_ folder creation should be skipped
        self.should_skip_wf_folder_creation = False

        # --- Fullscreen Toggle ---
        fullscreen_frame = ttk.Frame(self.master, padding="5 5 5 5")
        fullscreen_frame.pack(pady=5, fill="x")
        ttk.Checkbutton(fullscreen_frame, text="Maximize Window", variable=self.is_fullscreen,
                        command=self.toggle_fullscreen).pack(anchor="e")

        # Set initial window state (after a short delay to ensure window is ready)
        self.master.after(100, self.toggle_fullscreen)

        # --- Frame for Location Selection ---
        location_frame = ttk.Frame(self.master, padding="10 10 10 10")
        location_frame.pack(pady=5, fill="x")

        ttk.Label(location_frame, text="Selected Directory:").pack(anchor="w")
        # Initial wraplength, will be updated dynamically on window resize
        self.location_label = ttk.Label(location_frame, textvariable=self.output_location,
                                         wraplength=self.master.winfo_width() - 50, justify="left", foreground="blue")
        self.location_label.pack(fill="x", pady=(0, 5))
        self.output_location.set("No directory selected.")

        # "Browse for Location" button - now uses the specific "Browse.TButton" style
        ttk.Button(location_frame, text="Select The Main Folder", command=self.browse_location, style="Browse.TButton").pack(pady=5, fill="x")

        # --- Checkbutton for Auto-Pick Folder Name ---
        auto_pick_frame = ttk.Frame(self.master, padding="10 0 10 0")
        auto_pick_frame.pack(pady=(0, 5), fill="x")
        ttk.Checkbutton(auto_pick_frame, text="If you want to add some books select this option",
                        variable=self.auto_pick_code,
                        command=self._toggle_automatic_code).pack(anchor="w")

        # --- Frame for Main Folder Code ---
        main_code_frame = ttk.Frame(self.master, padding="10 10 10 10")
        main_code_frame.pack(pady=5, fill="x")

        self.code_label = ttk.Label(main_code_frame, text="Enter 4-digit code for WF_ folder:")
        self.code_label.pack(anchor="w")
        self.code_entry = ttk.Entry(main_code_frame, textvariable=self.main_folder_code, width=10)
        self.code_entry.pack(fill="x", pady=(0, 5))
        # Bind Enter key to move focus to the next input field
        self.code_entry.bind("<Return>", lambda event: self.num_sub_entry.focus_set())

        # --- Frame for Number of Sub-Folders (Books) ---
        num_sub_frame = ttk.Frame(self.master, padding="10 10 10 10")
        num_sub_frame.pack(pady=5, fill="x")
        ttk.Label(num_sub_frame, text="Number of Books in this carton / Total Number of books").pack(anchor="w")
        self.num_sub_entry = ttk.Entry(num_sub_frame, textvariable=self.num_sub_folders, width=5)
        self.num_sub_entry.pack(fill="x", pady=(0, 5))
        self.num_sub_folders.set("2") # Default value for number of books
        # Bind to generate dynamic entries when number is entered or focus is lost
        self.num_sub_entry.bind("<Return>", self.generate_sub_folder_inputs)
        self.num_sub_entry.bind("<FocusOut>", self.generate_sub_folder_inputs)

        # --- Scrollable Area for Dynamic Sub-Folder Names (Books) ---
        self.canvas = tk.Canvas(self.master, borderwidth=0, background="#f0f0f0")
        self.canvas.pack(side="left", fill="both", expand=True, padx=10, pady=5)

        self.scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y", pady=5)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.dynamic_sub_folder_frame = ttk.Frame(self.canvas, padding="0 0 0 0")
        # Create a window on the canvas to hold the dynamic frame
        self.canvas_frame_id = self.canvas.create_window((0, 0), window=self.dynamic_sub_folder_frame, anchor="nw")

        # Bind events to update scroll region and canvas window size
        self.dynamic_sub_folder_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        # Bind mouse wheel for scrolling on the canvas
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel) # Windows/Linux
        self.canvas.bind_all("<Button-4>", self._on_mousewheel) # macOS scroll up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel) # macOS scroll down

        # --- Create Folders Button ---
        self.create_button = ttk.Button(self.master, text="Create Folders", command=self.create_folders_action)
        self.create_button.pack(pady=10, fill="x", padx=10) # Make button fill width and add padding

        # --- Status and Progress Bar ---
        self.status_label = ttk.Label(self.master, text="Ready", foreground="gray", font=("Inter", 9))
        self.status_label.pack(pady=(0, 5))

        self.progress_bar = ttk.Progressbar(self.master, orient="horizontal", length=400, mode="indeterminate")
        self.progress_bar.pack(pady=(0, 10), fill="x", padx=10) # Make progress bar fill width and add padding
        self.progress_bar.stop() # Ensure it's stopped initially


        # Generate initial sub-folder inputs based on default value
        self.generate_sub_folder_inputs()

        # Set initial state of auto-pick code option
        self._toggle_automatic_code()


    def toggle_fullscreen(self):
        """Toggles the window between maximized and normal state."""
        if self.is_fullscreen.get():
            self.master.state('zoomed')
        else:
            self.master.state('normal')
        # Update wraplength after state change
        self.master.update_idletasks() # Ensure window size is updated before getting width
        self.location_label.config(wraplength=self.master.winfo_width() - 50)


    def on_frame_configure(self, event):
        """Update the scroll region of the canvas when the inner frame's size changes."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        """Resize the canvas's internal window to match the canvas width."""
        self.canvas.itemconfig(self.canvas_frame_id, width=self.canvas.winfo_width())

    def _on_mousewheel(self, event):
        """Handles mouse wheel scrolling for the canvas."""
        if event.num == 5 or event.delta == -120: # Scroll down
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta == 120: # Scroll up
            self.canvas.yview_scroll(-1, "units")

    def _extract_code_from_folder_name(self, folder_name):
        """
        Extracts a 4-digit code from the beginning of a folder name,
        optionally preceded by 'WF_'.
        Returns the 4-digit code as a string, or an empty string if not found.
        """
        # Regex to match: optional "WF_", followed by exactly four digits, at the start of the string
        match = re.match(r"^(?:WF_)?(\d{4}).*", folder_name, re.IGNORECASE)
        if match:
            return match.group(1) # Return the captured 4 digits
        return ""

    def browse_location(self, event=None):
        """Opens a file dialog to select the base directory and potentially auto-pick code."""
        selected_directory = filedialog.askdirectory(title="Select the base directory for folder creation")
        if selected_directory:
            self.output_location.set(selected_directory)
            self.status_label.config(text="Directory selected.", foreground="gray")

            if self.auto_pick_code.get():
                outermost_folder_name = os.path.basename(selected_directory)
                extracted_code = self._extract_code_from_folder_name(outermost_folder_name)

                if extracted_code:
                    self.main_folder_code.set(extracted_code)
                    self.status_label.config(text=f"Directory selected. Code auto-extracted: {extracted_code}", foreground="gray")
                    self.should_skip_wf_folder_creation = True
                else:
                    self.main_folder_code.set("")
                    self.status_label.config(text="Selected directory name does not start with a valid 4-digit code. Please rename or uncheck auto-extract.", foreground="orange")
                    self.should_skip_wf_folder_creation = False
            else:
                self.should_skip_wf_folder_creation = False

            self.code_entry.focus_set()
        else:
            self.output_location.set("No directory selected.")
            self.status_label.config(text="No directory selected.", foreground="red")
            self.should_skip_wf_folder_creation = False


    def _toggle_automatic_code(self):
        """Enables or disables the manual code entry based on the auto-pick checkbox."""
        if self.auto_pick_code.get():
            self.code_entry.config(state=tk.DISABLED)
            self.code_label.config(foreground="grey")
            current_dir = self.output_location.get()
            if current_dir and current_dir != "No directory selected.":
                outermost_folder_name = os.path.basename(current_dir)
                extracted_code = self._extract_code_from_folder_name(outermost_folder_name)

                if extracted_code:
                    self.main_folder_code.set(extracted_code)
                    self.status_label.config(text=f"Code auto-extracted from selected directory: {extracted_code}", foreground="gray")
                    self.should_skip_wf_folder_creation = True
                else:
                    self.main_folder_code.set("")
                    self.status_label.config(text="Selected directory name does not start with a valid 4-digit code. Please rename or uncheck auto-extract.", foreground="orange")
                    self.should_skip_wf_folder_creation = False
            else:
                self.main_folder_code.set("")
                self.status_label.config(text="Select a directory to auto-extract code.", foreground="orange")
                self.should_skip_wf_folder_creation = False
        else:
            self.code_entry.config(state=tk.NORMAL)
            self.code_label.config(foreground="black")
            self.main_folder_code.set("")
            self.status_label.config(text="Enter 4-digit code manually.", foreground="gray")
            self.should_skip_wf_folder_creation = False


    def generate_sub_folder_inputs(self, event=None):
        """
        Dynamically creates input fields for sub-folder names (Books),
        their corresponding chapter counts, and chapter naming format.
        """
        # Clear existing entries
        for widget in self.dynamic_sub_folder_frame.winfo_children():
            widget.destroy()
        self.book_inputs = [] # Reset the list of book data

        try:
            num_subs = int(self.num_sub_folders.get())
            if num_subs <= 0:
                ttk.Label(self.dynamic_sub_folder_frame, text="Enter a positive number for books.").pack()
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
                self.status_label.config(text="Invalid number of books.", foreground="red")
                return
        except ValueError:
            ttk.Label(self.dynamic_sub_folder_frame, text="Invalid number. Please enter an integer.").pack()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            self.status_label.config(text="Invalid number format.", foreground="red")
            return

        ttk.Label(self.dynamic_sub_folder_frame, text="Enter details for each book:").pack(anchor="w", pady=(5,0))

        # Create new entry fields for each book and its chapters
        for i in range(num_subs):
            book_data = {}

            # Frame to hold all inputs for a single book (Name, Chapters, Format)
            book_container_frame = ttk.Frame(self.dynamic_sub_folder_frame, padding="5 0 0 10", relief="groove", borderwidth=1)
            book_container_frame.pack(fill="x", pady=5, padx=2)

            # Book Name Input
            name_row_frame = ttk.Frame(book_container_frame)
            name_row_frame.pack(fill="x", pady=(5,2))
            ttk.Label(name_row_frame, text=f"Book {i+1} Name:", width=15).pack(side="left", padx=(0,5))
            name_var = tk.StringVar()
            book_data['name_var'] = name_var
            name_entry = ttk.Entry(name_row_frame, textvariable=name_var)
            name_entry.pack(side="left", fill="x", expand=True)
            book_data['name_entry'] = name_entry

            # Chapter Count Input for this book
            chapters_row_frame = ttk.Frame(book_container_frame)
            chapters_row_frame.pack(fill="x", pady=(2,2))
            ttk.Label(chapters_row_frame, text=f"Chapters Count:", width=15).pack(side="left", padx=(0,5))
            chapters_var = tk.StringVar()
            book_data['chapters_var'] = chapters_var
            chapters_entry = ttk.Entry(chapters_row_frame, textvariable=chapters_var, width=5)
            chapters_entry.pack(side="left")
            book_data['chapters_entry'] = chapters_entry

            # Chapter Format Selection for this book
            format_row_frame = ttk.Frame(book_container_frame)
            format_row_frame.pack(fill="x", pady=(2,5))
            ttk.Label(format_row_frame, text=f"Chapter Format:", width=15).pack(side="left", padx=(0,5))
            format_var = tk.StringVar(value="Digits") # Default to Digits
            book_data['format_var'] = format_var
            format_combobox = ttk.Combobox(format_row_frame, textvariable=format_var,
                                           values=["Digits", "Words", "Null"],
                                           state="readonly")
            format_combobox.pack(side="left", fill="x", expand=True)
            book_data['format_combobox'] = format_combobox

            self.book_inputs.append(book_data) # Store the dictionary for this book

            # Set default names and chapter counts for convenience
            if i == 0: name_var.set("Book A"); chapters_var.set("5")
            elif i == 1: name_var.set("Book B"); chapters_var.set("3")
            elif i == 2: name_var.set("Book C"); chapters_var.set("7")
            else: name_var.set(f"Book {i+1}"); chapters_var.set("5")

            # Bind Enter key for navigation/action
            name_entry.bind("<Return>", lambda e, entry=chapters_entry: entry.focus_set())
            chapters_entry.bind("<Return>", lambda e, cb=format_combobox: cb.focus_set())
            # Ensure the combobox also moves focus on selection or Enter
            format_combobox.bind("<<ComboboxSelected>>", lambda e, idx=i: self._focus_next_book_input(idx))
            format_combobox.bind("<Return>", lambda e, idx=i: self._focus_next_book_input(idx))


        # Update the scroll region after adding all widgets
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        # Set focus to the first dynamically created book name entry if available
        if self.book_inputs:
            self.book_inputs[0]['name_entry'].focus_set()
            self.status_label.config(text=f"Ready for {num_subs} book details.", foreground="gray")


    def _focus_next_book_input(self, current_index):
        """Helper to move focus to the next book's name entry or to the final 'Create Folders' button."""
        if current_index < len(self.book_inputs) - 1:
            next_book_data = self.book_inputs[current_index + 1]
            next_entry = next_book_data['name_entry']
            next_entry.focus_set()

            self.canvas.update_idletasks()

            # Get the bounding box of the next book's container frame
            book_container_frame = next_entry.winfo_parent().winfo_parent()

            entry_frame_y_top_relative_to_scrollable_content = book_container_frame.winfo_y()
            entry_frame_y_bottom_relative_to_scrollable_content = book_container_frame.winfo_y() + book_container_frame.winfo_height()

            canvas_view_top = self.canvas.canvasy(0)
            canvas_view_bottom = self.canvas.canvasy(self.canvas.winfo_height())

            if entry_frame_y_bottom_relative_to_scrollable_content > canvas_view_bottom:
                scroll_amount = entry_frame_y_bottom_relative_to_scrollable_content - canvas_view_bottom + 10
                self.canvas.yview_scroll(int(scroll_amount), "pixels")
            elif entry_frame_y_top_relative_to_scrollable_content < canvas_view_top:
                scroll_amount = entry_frame_y_top_relative_to_scrollable_content - canvas_view_top - 10
                self.canvas.yview_scroll(int(scroll_amount), "pixels")

        else:
            self.create_button.focus_set()

    def _number_to_word(self, num):
        """
        Converts a number (1-300) to its word representation.
        Adds double space between tens and units words for numbers > 20.
        """
        if not isinstance(num, int) or num < 1 or num > 300:
            return str(num) # Fallback to string for unsupported range

        words_under_20 = {
            1: "One", 2: "Two", 3: "Three", 4: "Four", 5: "Five",
            6: "Six", 7: "Seven", 8: "Eight", 9: "Nine", 10: "Ten",
            11: "Eleven", 12: "Twelve", 13: "Thirteen", 14: "Fourteen", 15: "Fifteen",
            16: "Sixteen", 17: "Seventeen", 18: "Eighteen", 19: "Nineteen"
        }

        tens_words = {
            20: "Twenty", 30: "Thirty", 40: "Forty", 50: "Fifty",
            60: "Sixty", 70: "Seventy", 80: "Eighty", 90: "Ninety"
        }

        if num < 20:
            return words_under_20[num]
        elif num < 100:
            tens = (num // 10) * 10
            units = num % 10
            if units == 0:
                return tens_words[tens]
            else:
                # Apply double space for numbers after 20 (e.g., 21, 35)
                return f"{tens_words[tens]}  {words_under_20[units]}"
        elif num < 301:
            hundreds = num // 100
            remainder = num % 100

            hundreds_word = ""
            if hundreds == 1:
                hundreds_word = "One Hundred"
            elif hundreds == 2:
                hundreds_word = "Two Hundred"
            elif hundreds == 3:
                hundreds_word = "Three Hundred"

            if remainder == 0:
                return hundreds_word
            else:
                # The recursive call self._number_to_word(remainder) will apply the double space
                # if 'remainder' is between 21-99.
                return f"{hundreds_word} and {self._number_to_word(remainder)}"
        else:
            return str(num)


    def create_folders_action(self):
        """Validates inputs and triggers the folder creation."""
        base_directory = self.output_location.get()
        main_code = self.main_folder_code.get()

        # --- Validation ---
        if not base_directory or base_directory == "No directory selected.":
            messagebox.showerror("Input Error", "Please select a base directory.")
            self.status_label.config(text="Please select a base directory.", foreground="red")
            return

        if not main_code:
            messagebox.showerror("Input Error", "Please enter the 4-digit code or ensure it was auto-extracted successfully from the selected directory.")
            self.status_label.config(text="4-digit code is missing.", foreground="red")
            return

        if not main_code.isdigit() or len(main_code) != 4:
            messagebox.showerror("Input Error", "The main folder code must be exactly 4 digits. Please correct it, or if auto-extracting, ensure the selected directory name starts with 'WF_XXXX' or 'XXXX'.")
            self.status_label.config(text="Invalid 4-digit code format.", foreground="red")
            return

        try:
            num_books_expected = int(self.num_sub_folders.get())
            if num_books_expected <= 0:
                messagebox.showerror("Input Error", "Number of books must be a positive integer.")
                self.status_label.config(text="Number of books must be positive.", foreground="red")
                return
        except ValueError:
            messagebox.showerror("Input Error", "Number of books must be a valid integer.")
            self.status_label.config(text="Invalid number of books.", foreground="red")
            return

        # Collect and validate book names and chapter counts and formats
        book_data_for_creation = []
        for i, book_info in enumerate(self.book_inputs):
            book_name = book_info['name_var'].get().strip()
            chapter_count_str = book_info['chapters_var'].get().strip()
            chapter_format_selected = book_info['format_var'].get() # Get format for this specific book

            if not book_name:
                messagebox.showerror("Input Error", f"Please enter a name for Book {i+1}.")
                self.status_label.config(text=f"Book {i+1} name is missing.", foreground="red")
                return

            try:
                num_chapters_val = int(chapter_count_str)
                if num_chapters_val <= 0:
                    messagebox.showerror("Input Error", f"Number of chapters for Book '{book_name}' must be a positive integer.")
                    self.status_label.config(text=f"Invalid chapters for Book '{book_name}'.", foreground="red")
                    return

                # Warn if using words for numbers outside the supported range (1-300) AND if Words format is chosen
                if chapter_format_selected == "Words" and (num_chapters_val < 1 or num_chapters_val > 300):
                    response = messagebox.askyesno("Warning", f"Using 'Words' format for chapter numbers for '{book_name}' (count: {num_chapters_val}) might result in digits for higher numbers. Do you want to continue?")
                    if not response:
                        self.status_label.config(text="Chapter creation cancelled.", foreground="orange")
                        return

                book_data_for_creation.append({
                    'name': book_name,
                    'chapters': num_chapters_val,
                    'format': chapter_format_selected # Store the format choice
                })

            except ValueError:
                messagebox.showerror("Input Error", f"Number of chapters for Book '{book_name}' must be a valid integer.")
                self.status_label.config(text=f"Invalid chapter format for Book '{book_name}'.", foreground="red")
                return

        if len(book_data_for_creation) != num_books_expected:
            messagebox.showerror("Input Error", f"Mismatch: Expected {num_books_expected} books, but found {len(book_data_for_creation)} valid entries. Please ensure all book names and chapter counts are filled correctly.")
            self.status_label.config(text="Mismatch in book/chapter counts.", foreground="red")
            return

        skip_wf_folder_creation_for_this_run = self.should_skip_wf_folder_creation

        self.create_button.config(state=tk.DISABLED)
        self.progress_bar.start()
        self.status_label.config(text="Creating folders...", foreground="blue")
        self.master.update_idletasks()

        try:
            # --- DEBUG PRINT ---
            print(f"DEBUG: book_data_for_creation before calling _create_nested_folders_logic: {book_data_for_creation}")
            self._create_nested_folders_logic(base_directory, main_code, book_data_for_creation,
                                              skip_wf_folder_creation_for_this_run)
            messagebox.showinfo("Success", "Folders created successfully!")
            self.status_label.config(text="Folders created successfully!", foreground="green")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status_label.config(text=f"Error: {e}", foreground="red")
        finally:
            self.create_button.config(state=tk.NORMAL)
            self.progress_bar.stop()

    def _create_nested_folders_logic(self, base_directory, main_folder_code, book_data,
                                      skip_wf_folder=False):
        # --- DEBUG PRINT ---
        print(f"DEBUG: _create_nested_folders_logic received book_data: {book_data}")

        """
        Core logic to create the nested folder structure.
        WF_[MainFolderName]
        └── [BookName1]
            ├── WF_[Main 4-digit Code]_[Book Name]_Cover
            ├── WF_[Main 4-digit Code]_[Book Name]_Front Index
            ...
            ├── WF_[Main 4-digit Code]_[Book Name]_Chapter Null (1)_Null Name_ OR WF_[Main 4-digit Code]_[Book Name]_Chapter One_Null Name_
            ├── WF_[Main 4-digit Code]_[Book Name]_Chapter Null (2)_Null Name_ OR WF_[Main 4-digit Code]_[Book Name]_Chapter Two_Null Name_
            ...
        """
        fixed_sub_sub_folders = [
            "Cover", "Front Index", "Acknowledgements", "Note About On The Author",
            "Author's Note", "Prologue", "Epilogue", "Back Index", "About The Author",
            "Prelude", "Preface", "Index", "Bibliography", "Appendix", "Notes"
        ]

        self.status_label.config(text="Ensuring base directory exists...", foreground="blue")
        self.master.update_idletasks()

        os.makedirs(base_directory, exist_ok=True)
        print(f"Ensured base directory exists: {base_directory}\n")

        # Conditional WF_ folder creation
        if skip_wf_folder:
            wf_folder_path = base_directory
            self.status_label.config(text=f"Skipping WF_ folder creation. Using '{os.path.basename(base_directory)}' as base.", foreground="blue")
            print(f"Skipping WF_ folder creation. Using '{wf_folder_path}' as the root for sub-folders.")
        else:
            self.status_label.config(text=f"Creating WF_{main_folder_code}...", foreground="blue")
            self.master.update_idletasks()

            wf_folder_name = f"WF_{main_folder_code}"
            wf_folder_path = os.path.join(base_directory, wf_folder_name)

            os.makedirs(wf_folder_path, exist_ok=True)
            print(f"Created main WF_ folder: {wf_folder_path}")

        # Create sub-folders (Books) directly inside the determined wf_folder_path
        for i, book_info in enumerate(book_data):
            sub_name_raw = book_info['name']
            num_chapters_for_this_book = book_info['chapters']
            chapter_format_for_this_book = book_info['format'] # Get the format for this specific book

            sub_name_final = f"WF_{main_folder_code}_{sub_name_raw}"

            self.status_label.config(text=f"Creating book folder {i+1}/{len(book_data)}: {sub_name_final}...", foreground="blue")
            self.master.update_idletasks()

            sub_folder_path = os.path.join(wf_folder_path, sub_name_final)
            os.makedirs(sub_folder_path, exist_ok=True)
            print(f"  Created book folder: {sub_folder_path}")

            # --- Create fixed sub-sub-folders inside each book's folder ---
            for j, fixed_name_raw in enumerate(fixed_sub_sub_folders):
                fixed_folder_name_final = f"WF_{main_folder_code}_{sub_name_raw}_{fixed_name_raw}"
                self.status_label.config(text=f"    Creating fixed folder {j+1}/{len(fixed_sub_sub_folders)} in '{sub_name_final}': {fixed_folder_name_final}...", foreground="darkgreen")
                self.master.update_idletasks()

                fixed_folder_path = os.path.join(sub_folder_path, fixed_folder_name_final)
                os.makedirs(fixed_folder_path, exist_ok=True)
                print(f"    Created fixed sub-sub-folder: {fixed_folder_path}")
            print(f"  Finished fixed folders for '{sub_name_final}'")
            # --- End fixed sub-sub-folder creation ---

            # --- Create Chapter folders at the same level as fixed folders ---
            for k in range(1, num_chapters_for_this_book + 1):
                chapter_folder_name_final = ""
                status_text_chapter_progress = ""
                null_name_suffix = "_Null Name" # Define the suffix once

                if chapter_format_for_this_book == "Null":
                    chapter_folder_name_final = f"WF_{main_folder_code}_{sub_name_raw}_Chapter Null ({k}){null_name_suffix}"
                    status_text_chapter_progress = f"    Creating Chapter Null folder {k}/{num_chapters_for_this_book} in '{sub_name_final}': {chapter_folder_name_final}..."
                elif chapter_format_for_this_book == "Words":
                    chapter_num_str = self._number_to_word(k)
                    chapter_folder_name_final = f"WF_{main_folder_code}_{sub_name_raw}_Chapter {chapter_num_str}{null_name_suffix}"
                    status_text_chapter_progress = f"    Creating Chapter Word folder {k}/{num_chapters_for_this_book} in '{sub_name_final}': {chapter_folder_name_final}..."
                else: # "Digits" (default)
                    chapter_num_str = str(k)
                    chapter_folder_name_final = f"WF_{main_folder_code}_{sub_name_raw}_Chapter {chapter_num_str}{null_name_suffix}"
                    status_text_chapter_progress = f"    Creating Chapter Digit folder {k}/{num_chapters_for_this_book} in '{sub_name_final}': {chapter_folder_name_final}..."

                self.status_label.config(text=status_text_chapter_progress, foreground="darkblue")
                self.master.update_idletasks()
                chapter_folder_path = os.path.join(sub_folder_path, chapter_folder_name_final)
                os.makedirs(chapter_folder_path, exist_ok=True)
                print(f"    Created chapter folder: {chapter_folder_path}")
            print(f"  Finished chapters for '{sub_name_final}'")
            # --- End Chapter folder creation ---
        print("-" * 30)
        print("\nFolder creation process completed.")

# --- Run the App ---
if __name__ == "__main__":
    root = tk.Tk()
    app = FolderCreatorApp(root)
    root.mainloop()
