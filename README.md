# Tax Document Merger 📄

A Python script designed to merge employee tax certificates with their corresponding challan documents into a single PDF file, guided by an Excel mapping file.

## Features

* **Automated Merging:** Automatically combines multiple challan PDFs with a master tax certificate.
* **Excel-Driven:** Uses a simple Excel sheet to define the mapping between certificates and challans.
* **User-Friendly:** A graphical user interface (GUI) prompts you to select the correct folders and files.

## Getting Started

### Prerequisites

Before running the script, you must set up the following directory structure:

your-project-folder/


├── Certificate/


│   └── (all your tax certificate PDFs go here)


├── Challan/


│   └── (all your challan PDFs go here)


└── Output/


└── (merged PDFs will be saved here)

You also need an **Excel mapping file**. This file acts as the brain for the operation, telling the script which files to combine. The Excel file should contain columns that specify the certificate filename and the challan filename(s) associated with it.

### How to Run

1.  **Organize Files:** Place all tax certificates in the `Certificate` folder and all challans in the `Challan` folder.
2.  **Execute the Script:** Run the Python file or executable file.
3.  **Select Directories:**
4.  **Done!** The script will automatically process the files based on your Excel sheet and save the merged PDFs into the `Output` folder.

## Workflow ⚙️

The script performs the following steps:

1.  **Prompts for Input:** Asks the user to specify the main working directory and the location of the Excel file.
2.  **Reads Mapping:** Opens the Excel file to understand which challans should be merged with each tax certificate.
3.  **Merges PDFs:** For each entry, it finds the specified certificate and challan(s), merges them in order, and creates a new, combined PDF.
4.  **Saves Output:** Saves the newly created PDF files in the `Output` directory.
