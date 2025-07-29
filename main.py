"""
Excel Worksheet Unprotector

A desktop application to remove worksheet password protection
from .xlsx files.

Authors:
    - Amirhossein Noshadi <amirhosseinnoshadi@cnr.it>
    - Enrico Cambiaso <enrico.cambiaso@cnr.it>
    
Email : <amirhosseinnoshadi@cnr.it> & <enrico.cambiaso@cnr.it>
Copyright (c) 2025 Amirhossein Noshadi & Enrico Cambiaso
License: MIT License
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import re
import os
import threading

def remove_excel_protection(input_filename, output_filename, status_callback):
    """
    Removes worksheet protection from an .xlsx file using string replacement
    to avoid XML parsing errors and file corruption.
    
    Args:
        input_filename (str): Path to the protected .xlsx file.
        output_filename (str): Path to save the unprotected .xlsx file.
        status_callback (function): Function to send status updates to the GUI.
    """
    try:
        # Regular expression to find the <sheetProtection ... /> tag.
        # This is more robust than simple string finding.
        protection_regex = re.compile(r"<sheetProtection.*?/>")
        protection_found_in_any_sheet = False

        with zipfile.ZipFile(input_filename, 'r') as zin:
            with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    # Read the original file content
                    buffer = zin.read(item.filename)

                    # Only process worksheet XML files
                    if item.filename.startswith('xl/worksheets/'):
                        # Decode the XML file to a string for processing
                        xml_content = buffer.decode('utf-8')
                        
                        # Use the regex to find and remove the protection tag
                        new_xml_content, num_replacements = protection_regex.subn('', xml_content)
                        
                        if num_replacements > 0:
                            protection_found_in_any_sheet = True
                            status_callback(f"Found and removed protection in: {item.filename}")
                            # Write the modified content back to the new zip
                            zout.writestr(item, new_xml_content.encode('utf-8'))
                        else:
                            # If no protection was found, write the original content back
                            zout.writestr(item, buffer)
                    else:
                        # For all other files, just copy them to the new zip
                        zout.writestr(item, buffer)
        
        if protection_found_in_any_sheet:
            status_callback(f"Success! Unprotected file saved as:\n{output_filename}")
            messagebox.showinfo("Success", f"File has been unprotected and saved successfully!")
        else:
            status_callback("No worksheet protection was found in the file.")
            messagebox.showwarning("No Protection Found", "Could not find any worksheet protection in the selected file.")

    except Exception as e:
        status_callback(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred during processing:\n{e}")

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Worksheet Unprotector")
        self.root.geometry("500x350")
        self.root.resizable(False, False)
        
        self.input_file_path = ""

        # --- UI Elements ---
        self.main_frame = tk.Frame(root, padx=20, pady=20)
        self.main_frame.pack(fill="both", expand=True)

        self.select_button = tk.Button(self.main_frame, text="1. Select Excel File (.xlsx)", command=self.select_file, font=("Helvetica", 12))
        self.select_button.pack(pady=10, fill="x")

        self.file_label = tk.Label(self.main_frame, text="No file selected", wraplength=450, fg="gray", font=("Helvetica", 10))
        self.file_label.pack(pady=5)

        self.run_button = tk.Button(self.main_frame, text="2. Remove Protection and Save", command=self.run_process_thread, state="disabled", font=("Helvetica", 12, "bold"))
        self.run_button.pack(pady=10, fill="x")

        self.status_label = tk.Label(self.main_frame, text="Status: Waiting for file...", wraplength=450, justify="left", font=("Helvetica", 10))
        self.status_label.pack(pady=10, fill="x")

    def select_file(self):
        """Opens a file dialog to select an .xlsx file."""
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.input_file_path = path
            self.file_label.config(text=os.path.basename(path), fg="black")
            self.status_label.config(text=f"Status: Ready to process '{os.path.basename(path)}'")
            self.run_button.config(state="normal")
        else:
            self.input_file_path = ""
            self.file_label.config(text="No file selected", fg="gray")
            self.status_label.config(text="Status: Waiting for file...")
            self.run_button.config(state="disabled")

    def update_status(self, message):
        """Callback to update the status label from the processing thread."""
        self.status_label.config(text=f"Status: {message}")

    def run_process_thread(self):
        """Runs the main logic in a separate thread to keep the GUI responsive."""
        if not self.input_file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=os.path.basename(self.input_file_path).replace(".xlsx", "_unprotected.xlsx")
        )

        if output_path:
            self.run_button.config(state="disabled")
            self.select_button.config(state="disabled")
            self.update_status("Processing... This may take a moment.")
            
            thread = threading.Thread(
                target=self.run_process_logic,
                args=(output_path,)
            )
            thread.start()

    def run_process_logic(self, output_path):
        """The actual logic to be run in the thread."""
        remove_excel_protection(self.input_file_path, output_path, self.update_status)
        # Re-enable buttons once done
        self.run_button.config(state="normal")
        self.select_button.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

