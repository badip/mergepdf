# PDF Merger Script with GUI and Logging
# This script provides a graphical interface to merge certificate PDFs with corresponding challan PDFs.
# It uses an Excel file to map employees to their challan numbers and provides detailed logging.

# --- Prerequisites ---
# Before running, you need to install the required Python libraries.
# You can install them by opening your terminal or command prompt and running:
# pip install pandas openpyxl pypdf

import os
import pandas as pd
from pypdf import PdfWriter
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import logging
from logging.handlers import QueueHandler
import queue
import threading
import sys

# --- Set up Logging ---
# This function configures logging to go to both a file and the GUI.
def setup_logging(log_queue):
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    
    # Create a root logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # Prevent duplicate handlers if this function is called again
    if logger.hasHandlers():
        logger.handlers.clear()
        
    # 1. File Handler: Saves logs to a file for debugging crashes.
    file_handler = logging.FileHandler('pdf_merger.log', mode='w')
    file_handler.setFormatter(logging.Formatter(log_format))
    logger.addHandler(file_handler)
    
    # 2. Queue Handler: Sends logs to the GUI.
    queue_handler = QueueHandler(log_queue)
    queue_handler.setFormatter(logging.Formatter(log_format))
    logger.addHandler(queue_handler)

    # Redirect stdout and stderr to the logger
    # This ensures that any unexpected errors or prints are also captured.
    sys.stdout = LogRedirector(logging.INFO)
    sys.stderr = LogRedirector(logging.ERROR)

class LogRedirector:
    """A class to redirect stdout/stderr to the logging module."""
    def __init__(self, level):
        self.level = level

    def write(self, message):
        if message.rstrip() != "":
            logging.log(self.level, message.rstrip())

    def flush(self):
        pass

# --- Main Application Logic ---
def merge_pdfs_logic(paths):
    """The core logic for finding and merging PDFs."""
    cert_dir, challan_dir, output_dir, excel_file = paths
    
    logging.info("Starting the PDF merging process...")
    logging.info(f"Certificate Directory: {cert_dir}")
    logging.info(f"Challan Directory: {challan_dir}")
    logging.info(f"Output Directory: {output_dir}")
    logging.info(f"Excel File: {excel_file}")

    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        df.columns = [str(col).strip() for col in df.columns]
        employee_col_name = 'Employee Name'
        challan_col_name = 'Challan Number'

        if employee_col_name not in df.columns or challan_col_name not in df.columns:
            logging.error(f"Excel file must contain columns named '{employee_col_name}' and '{challan_col_name}'.")
            logging.error(f"Found columns: {df.columns.tolist()}")
            return

        logging.info("Successfully loaded and validated the Excel file.")

    except FileNotFoundError:
        logging.error(f"The Excel file was not found at {excel_file}")
        return
    except Exception as e:
        logging.error(f"An error occurred while reading the Excel file: {e}", exc_info=True)
        return

    processed_files = 0
    total_certs = [f for f in os.listdir(cert_dir) if f.lower().endswith('.pdf')]
    logging.info(f"Found {len(total_certs)} PDF files in the certificate directory.")

    for cert_filename in total_certs:
        employee_name = os.path.splitext(cert_filename)[0]
        logging.info(f"--- Processing certificate for: {employee_name} ---")

        try:
            employee_challans = df[df[employee_col_name].astype(str).str.strip() == employee_name.strip()]

            if employee_challans.empty:
                logging.warning(f"No challan entries found for '{employee_name}' in the Excel file. Skipping.")
                continue

            merger = PdfWriter()
            
            cert_path = os.path.join(cert_dir, cert_filename)
            merger.append(cert_path)
            logging.info(f"Added certificate: {cert_filename}")

            for _, row in employee_challans.iterrows():
                challan_num = str(row[challan_col_name]).strip()
                challan_filename = f"{challan_num}.pdf"
                challan_path = os.path.join(challan_dir, challan_filename)

                if os.path.exists(challan_path):
                    try:
                        merger.append(challan_path)
                        logging.info(f"  + Added challan: {challan_filename}")
                    except Exception as e:
                        logging.warning(f"  - Could not merge challan file {challan_path}. Skipping. Error: {e}")
                else:
                    logging.warning(f"  - Challan file not found: {challan_path}. Skipping.")
            
            output_filename = f"{employee_name}_combined.pdf"
            output_path = os.path.join(output_dir, output_filename)
            
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            merger.close()
            logging.info(f"Successfully created merged file: {output_path}")
            processed_files += 1

        except Exception as e:
            logging.error(f"A critical error occurred while processing {employee_name}. Skipping. Error: {e}", exc_info=True)
            continue
            
    logging.info("--- Process Complete ---")
    logging.info(f"Successfully processed and merged PDFs for {processed_files} out of {len(total_certs)} employees.")


# --- GUI Class ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Merger")
        self.geometry("800x600")
        self.paths = {"cert": tk.StringVar(), "challan": tk.StringVar(), "output": tk.StringVar(), "excel": tk.StringVar()}

        # Set up the main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Path Selection UI ---
        controls_frame = ttk.LabelFrame(main_frame, text="Setup", padding="10")
        controls_frame.pack(fill=tk.X, expand=False, pady=5)
        controls_frame.grid_columnconfigure(1, weight=1)

        # Create rows for each path selection
        self.create_path_row(controls_frame, "Certificate Directory:", self.paths["cert"], 0, self.select_directory)
        self.create_path_row(controls_frame, "Challan Directory:", self.paths["challan"], 1, self.select_directory)
        self.create_path_row(controls_frame, "Output Directory:", self.paths["output"], 2, self.select_directory)
        self.create_path_row(controls_frame, "Excel File:", self.paths["excel"], 3, self.select_file)

        # --- Action Button ---
        self.start_button = ttk.Button(main_frame, text="Start Merging", command=self.start_processing)
        self.start_button.pack(pady=10, fill=tk.X)

        # --- Logging Text Area ---
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled', wrap=tk.WORD, bg="#f0f0f0", fg="black")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # --- Setup Logging Queue ---
        self.log_queue = queue.Queue()
        setup_logging(self.log_queue)
        self.after(100, self.process_log_queue)

    def create_path_row(self, parent, label_text, string_var, row, command):
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", padx=5, pady=5)
        entry = ttk.Entry(parent, textvariable=string_var, state="readonly")
        entry.grid(row=row, column=1, sticky="ew", padx=5)
        button = ttk.Button(parent, text="Browse...", command=lambda: command(string_var))
        button.grid(row=row, column=2, sticky="e", padx=5)

    def select_directory(self, string_var):
        path = filedialog.askdirectory()
        if path:
            string_var.set(path)

    def select_file(self, string_var):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            string_var.set(path)

    def start_processing(self):
        # Validate that all paths have been selected
        if not all(var.get() for var in self.paths.values()):
            logging.error("All paths must be selected before starting.")
            return

        self.start_button.config(state="disabled", text="Processing...")
        
        # Run the merging logic in a separate thread to keep the GUI responsive
        paths_tuple = (self.paths["cert"].get(), self.paths["challan"].get(), self.paths["output"].get(), self.paths["excel"].get())
        processing_thread = threading.Thread(target=self.run_merger_thread, args=(paths_tuple,), daemon=True)
        processing_thread.start()

    def run_merger_thread(self, paths_tuple):
        try:
            merge_pdfs_logic(paths_tuple)
        except Exception as e:
            logging.critical(f"An unhandled exception occurred in the processing thread: {e}", exc_info=True)
        finally:
            # Re-enable the button once processing is complete
            self.start_button.config(state="normal", text="Start Merging")

    def process_log_queue(self):
        """Checks the queue for new log messages and updates the GUI."""
        while not self.log_queue.empty():
            record = self.log_queue.get(block=False)
            msg = logging.getLogger().handlers[1].formatter.format(record) # Get formatter from QueueHandler
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, msg + '\n')
            self.log_text.config(state='disabled')
            self.log_text.yview(tk.END)
        self.after(100, self.process_log_queue)


if __name__ == "__main__":
    app = App()
    app.mainloop()
