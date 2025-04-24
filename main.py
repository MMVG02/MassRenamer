import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import customtkinter as ctk
import pandas as pd
import os
import re # Import regular expressions for natural sorting

# --- Constants ---
APP_NAME = "Mass Renamer"
WINDOW_WIDTH = 450
WINDOW_HEIGHT = 350

COLOR_BTN_XLSX = "#2ECC71" # Green
COLOR_BTN_FOLDER = "#3498DB" # Blue
COLOR_BTN_START = "#E74C3C" # Red
COLOR_BTN_HOVER = "#555555" # Dark Gray for hover

# --- Helper Function for Natural Sorting ---
def natural_sort_key(s):
    """
    Generate a key for sorting strings in natural order (alphanumeric).
    Example: ["item1", "item10", "item2"] -> ["item1", "item2", "item10"]
    """
    # Split the string into alternating non-digit and digit sequences.
    # Convert digit sequences to integers for proper numerical comparison.
    # Keep non-digit sequences as lowercase strings for case-insensitive text comparison.
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split('(\d+)', str(s))] # Ensure input is string

# --- Main Application Class ---
class MassRenamerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.xlsx_path = None
        self.folder_path = None

        # --- Window Setup ---
        self.title(APP_NAME)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        ctk.set_appearance_mode("Light") # Force Light Mode for white UI
        ctk.set_default_color_theme("blue") # Default theme

        # Center the window on launch
        self.center_window()

        # --- UI Elements ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Title Label
        self.title_label = ctk.CTkLabel(self.main_frame, text=APP_NAME, font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(pady=(0, 20))

        # --- XLSX Selection ---
        self.xlsx_frame = ctk.CTkFrame(self.main_frame)
        self.xlsx_frame.pack(fill="x", pady=5)

        self.xlsx_button = ctk.CTkButton(
            self.xlsx_frame,
            text="Select XLSX File",
            command=self.select_xlsx,
            fg_color=COLOR_BTN_XLSX,
            hover_color=COLOR_BTN_HOVER
        )
        self.xlsx_button.pack(side="left", padx=(0, 10))

        self.xlsx_label = ctk.CTkLabel(self.xlsx_frame, text="No file selected", text_color="gray", anchor="w")
        self.xlsx_label.pack(side="left", fill="x", expand=True)

        # --- Folder Selection ---
        self.folder_frame = ctk.CTkFrame(self.main_frame)
        self.folder_frame.pack(fill="x", pady=5)

        self.folder_button = ctk.CTkButton(
            self.folder_frame,
            text="Select Folder",
            command=self.select_folder,
            fg_color=COLOR_BTN_FOLDER,
            hover_color=COLOR_BTN_HOVER
        )
        self.folder_button.pack(side="left", padx=(0, 10))

        self.folder_label = ctk.CTkLabel(self.folder_frame, text="No folder selected", text_color="gray", anchor="w")
        self.folder_label.pack(side="left", fill="x", expand=True)

        # --- Start Button ---
        self.start_button = ctk.CTkButton(
            self.main_frame,
            text="Start!",
            command=self.start_renaming,
            fg_color=COLOR_BTN_START,
            hover_color=COLOR_BTN_HOVER,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.start_button.pack(pady=25, ipady=5, fill='x')

        # --- Status Label ---
        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="gray", anchor="center")
        self.status_label.pack(pady=(10, 0), fill='x')


    def center_window(self):
        """Centers the window on the screen."""
        self.update_idletasks() # Ensure geometry is calculated
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (WINDOW_WIDTH // 2)
        y = (screen_height // 2) - (WINDOW_HEIGHT // 2)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}")

    def select_xlsx(self):
        """Opens a dialog to select an XLSX file."""
        file_path = fd.askopenfilename(
            title="Select XLSX File",
            filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
        )
        if file_path:
            self.xlsx_path = file_path
            filename = os.path.basename(file_path)
            max_len = 35
            display_name = (filename[:max_len] + '...') if len(filename) > max_len else filename
            self.xlsx_label.configure(text=display_name, text_color="black")
            self.status_label.configure(text="")
        else:
            self.xlsx_path = None
            self.xlsx_label.configure(text="No file selected", text_color="gray")

    def select_folder(self):
        """Opens a dialog to select a folder."""
        folder_path = fd.askdirectory(title="Select Folder Containing Files to Rename")
        if folder_path:
            self.folder_path = folder_path
            max_len = 35
            display_name = (folder_path[:max_len] + '...') if len(folder_path) > max_len else folder_path
            self.folder_label.configure(text=display_name, text_color="black")
            self.status_label.configure(text="")
        else:
            self.folder_path = None
            self.folder_label.configure(text="No folder selected", text_color="gray")

    def start_renaming(self):
        """Performs the file renaming process."""
        self.status_label.configure(text="Processing...", text_color="orange")
        self.update_idletasks()

        if not self.xlsx_path or not self.folder_path:
            mb.showerror("Error", "Please select both an XLSX file and a Folder first.")
            self.status_label.configure(text="Error: Selection missing.", text_color="red")
            return

        warning_message = ""
        renamed_count = 0

        try:
            # 1. Read XLSX file
            try:
                df = pd.read_excel(self.xlsx_path, header=None)
                if df.empty or df.shape[1] == 0:
                     mb.showerror("Error", "The selected XLSX file is empty or has no columns.")
                     self.status_label.configure(text="Error: XLSX empty.", text_color="red")
                     return
                # Ensure names are read as strings to prevent pandas from inferring types
                new_names = df.iloc[:, 0].astype(str).tolist()
            except FileNotFoundError:
                 mb.showerror("Error", f"XLSX file not found:\n{self.xlsx_path}")
                 self.status_label.configure(text="Error: XLSX not found.", text_color="red")
                 return
            except Exception as e:
                 mb.showerror("Error", f"Failed to read XLSX file:\n{e}")
                 self.status_label.configure(text="Error: Cannot read XLSX.", text_color="red")
                 return

            # 2. List files in the folder
            try:
                all_entries = os.listdir(self.folder_path)
                old_files = [f for f in all_entries if os.path.isfile(os.path.join(self.folder_path, f))]
                # --- THIS IS THE KEY CHANGE ---
                # Sort files using the natural sort key
                old_files.sort(key=natural_sort_key)
                # --- END OF KEY CHANGE ---
            except FileNotFoundError:
                 mb.showerror("Error", f"Folder not found:\n{self.folder_path}")
                 self.status_label.configure(text="Error: Folder not found.", text_color="red")
                 return
            except Exception as e:
                 mb.showerror("Error", f"Failed to list files in folder:\n{e}")
                 self.status_label.configure(text="Error: Cannot access folder.", text_color="red")
                 return

            # 3. Check for count mismatch
            if len(old_files) != len(new_names):
                warning_message = "Warning: The number of files and names are not the same.\n"
                mb.showwarning("Warning", warning_message + "Only matching pairs will be renamed.")

            # 4. Determine how many files to rename
            num_to_rename = min(len(old_files), len(new_names))
            if num_to_rename == 0:
                mb.showinfo("Info", "No files to rename (either folder or XLSX list is empty).")
                self.status_label.configure(text="Nothing to rename.", text_color="gray")
                return

            # 5. Perform renaming
            skipped_due_to_conflict = 0
            errors_during_rename = 0
            for i in range(num_to_rename):
                old_name = old_files[i]
                new_name = new_names[i].strip() # Remove leading/trailing whitespace from Excel name

                # Basic validation for the new name
                if not new_name:
                    print(f"Skipping empty new name from Excel row {i+1}") # Log to console
                    errors_during_rename += 1
                    continue
                # Check for invalid characters (you might want to expand this list)
                # Common invalid chars in Windows filenames: <>:"/\|?*
                if any(char in new_name for char in '<>:"/\\|?*'):
                     print(f"Skipping invalid new name '{new_name}' (contains forbidden characters) from Excel row {i+1}")
                     errors_during_rename += 1
                     continue

                old_path = os.path.join(self.folder_path, old_name)
                new_path = os.path.join(self.folder_path, new_name)

                # Skip if old and new names are the same after potential path normalization
                if os.path.normpath(old_path) == os.path.normpath(new_path):
                    renamed_count += 1 # Count it as "done"
                    continue

                # Prevent overwriting existing files UNLESS it's the source file itself
                if os.path.exists(new_path):
                     print(f"Skipping rename: Target file '{new_name}' already exists.")
                     skipped_due_to_conflict += 1
                     continue

                try:
                    print(f"Renaming: '{old_name}' -> '{new_name}'") # Optional: Log rename action
                    os.rename(old_path, new_path)
                    renamed_count += 1
                except OSError as e:
                    errors_during_rename += 1
                    print(f"Error renaming '{old_name}' to '{new_name}': {e}") # Log error
                    if errors_during_rename == 1: # Show only the first error to avoid spamming
                         mb.showerror("Renaming Error", f"Could not rename '{old_name}' to '{new_name}'.\nError: {e}\n\nCheck file permissions or if the file is open.")


            # 6. Show final results
            success_message = f"{renamed_count} files renamed!"
            final_message = warning_message + success_message
            if skipped_due_to_conflict > 0:
                final_message += f"\n({skipped_due_to_conflict} skipped due to existing target name)"
            if errors_during_rename > 0:
                 final_message += f"\n({errors_during_rename} errors occurred during rename - see console/log)"


            mb.showinfo("Finished", final_message)
            self.status_label.configure(text=success_message, text_color="green")

        except Exception as e:
            mb.showerror("Unexpected Error", f"An unexpected error occurred: {e}")
            self.status_label.configure(text="Unexpected Error.", text_color="red")
            print(f"Unexpected Error: {e}")

# --- Run the Application ---
if __name__ == "__main__":
    app = MassRenamerApp()
    app.mainloop()
