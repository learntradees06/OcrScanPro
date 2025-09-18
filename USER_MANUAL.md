# OcrScanPro User Manual

## 1. Overview

OcrScanPro is a desktop application designed to help you extract key information from images of hardware components. Using Optical Character Recognition (OCR), it can identify and pull out data such as Part Numbers (PN), Manufacturer Part Numbers (MPN), Customer Part Numbers (CPN), and Serial Numbers (SN).

The application also features a local database to store mappings between MPNs and CPNs, which can be managed directly through the application's graphical user interface (GUI).

## 2. Installation and Setup

To run OcrScanPro, you need to have Python installed on your system, along with a few dependencies.

### 2.1. Prerequisites

*   **Python 3:** Make sure you have Python 3 installed. You can download it from [python.org](https://python.org).
*   **Tesseract OCR Engine:** This application uses the Tesseract engine for OCR. You must install it on your system.
    *   **Windows:** Download and run the installer from the [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) page. Make sure to note the installation path. The application will try to find it automatically, but if it fails, you may need to set the path manually in the `main.py` script.
    *   **macOS:** You can install it using Homebrew: `brew install tesseract`
    *   **Linux (Debian/Ubuntu):** You can install it using apt: `sudo apt-get update && sudo apt-get install -y tesseract-ocr`

### 2.2. Install Python Dependencies

Once the prerequisites are met, you can install the required Python libraries. Navigate to the project directory in your terminal and run the following command:

```bash
pip install -r requirements.txt
```

This will install all necessary libraries, including `opencv`, `pytesseract`, `pandas`, and `openpyxl`.

## 3. Running the Application

To start the application, run the `main.py` script from your terminal:

```bash
python main.py
```

This will launch the main application window.

## 4. Core Features

### 4.1. Extracting Data from an Image

You can extract data from an image in two ways:

*   **Upload Picture:** Click the **"1. Upload Picture and Extract Data"** button. This will open a file dialog. Select an image file (`.png`, `.jpg`, etc.) from your computer. The application will immediately process the image and display the results.
*   **Camera Capture:** Click the **"2. Camera Capture and Extract Data"** button. This will open your computer's webcam in a new window. Position the hardware component in front of the camera and click the "Capture Image" button. The captured image will then be processed.

The extracted data will appear in the "Extraction Results" text box, and the results will also be automatically saved to a file named `extracted_data.xlsx` in the project directory.

### 4.2. Using the "Deep Scan" Feature

For images that are difficult to read (e.g., blurry, poor lighting, unusual fonts), you can use the "Deep Scan" feature.

*   Before uploading or capturing an image, check the **"Enable Deep Scan (Slower)"** checkbox.
*   With this option enabled, the application will try multiple rotations (0, 90, 180, 270 degrees) and apply different image enhancement techniques to improve the chances of a successful extraction.
*   **Note:** As the name implies, this process is significantly slower than a standard scan.

## 5. Database Management

The application uses a local SQLite database (`cpn_database.db`) to store a mapping of Manufacturer Part Numbers (MPNs) to your internal Customer Part Numbers (CPNs). When an MPN is extracted from an image, the application will automatically look it up in this database and populate the CPN field with the result.

To manage this database, click the **"3. Manage Database"** button on the main window. This will open the Database Management window.

### 5.1. Viewing and Selecting Records

The main part of the Database Management window is a table that displays all the MPN/CPN pairs currently in the database. You can click on any row in this table to select it, which will populate the "Manage Record" fields below for easy editing or deletion.

### 5.2. Adding, Updating, and Deleting Records

You can manage records using the buttons provided:

*   **Add:** To add a new record, type the MPN and CPN into their respective entry fields and click the "Add" button.
*   **Update:** To update an existing record, first click on the record in the table to select it. Then, modify the values in the entry fields and click the "Update" button.
*   **Delete:** To delete a record, select it from the table and click the "Delete" button. You will be asked to confirm the deletion.
*   **Clear Fields:** This button will clear the text from the MPN and CPN entry fields.

### 5.3. Importing from Excel

This feature allows you to bulk-add records to the database from an Excel file.

1.  Click the **"Import from Excel"** button.
2.  Select an Excel file (`.xlsx`) from your computer.
3.  **Important:** The Excel file must contain two columns with the exact headers `mpn` and `cpn` (in lowercase).
4.  The application will read the file and attempt to insert each row into the database.
5.  If an MPN from the Excel file already exists in the database, that row will be **skipped** to prevent overwriting your data.
6.  A summary message will appear, telling you how many records were successfully imported and how many were skipped.

### 5.4. Exporting to Excel

This feature allows you to back up or share your database.

1.  Click the **"Export to Excel"** button.
2.  A "Save As" dialog will appear. Choose a location and a name for your export file.
3.  The application will save all records from the database into the specified Excel file.
4.  A confirmation message will appear when the export is complete.
