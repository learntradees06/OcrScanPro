# Hardware Component OCR Data Extractor - User Manual

## 1. Introduction

Welcome to the Hardware Component OCR Data Extractor application! This tool is designed to help you quickly extract key information from images of hardware components, such as Part Numbers (PN), Manufacturer Part Numbers (MPN), and Serial Numbers (SN).

The application features three main sections, organized into tabs:
- **OCR Extraction:** The main tab for extracting data from images.
- **Database Management:** A section to manage the mapping between Manufacturer Part Numbers (MPN) and your internal Customer Part Numbers (CPN).
- **Model Training:** A powerful feature that allows you to label your own images to improve the accuracy of the OCR engine over time.

## 2. OCR Extraction Tab

This is where the core data extraction happens.

### How to Extract Data:

1.  **Choose your source:**
    *   **Upload Picture & Extract:** Click this button to open a file dialog. Select an image file (`.png`, `.jpg`, etc.) from your computer.
    *   **Camera Capture & Extract:** Click this button to open your computer's camera. A new window will appear with the live camera feed.
        *   Point the camera at the hardware component.
        *   Click the **Capture Image** button to take a picture.
        *   The camera window will close, and the captured image will be processed.

2.  **Enable Deep Scan (Optional):**
    *   Before starting the extraction, you can check the **Enable Deep Scan (Slower)** box.
    *   This option tells the application to rotate the image in 90-degree increments and apply different preprocessing techniques. It is more thorough and can find text on poorly oriented images, but it takes significantly more time.

3.  **View the Results:**
    *   Once you select an image or capture a photo, the process will start automatically. You can monitor the progress via the status bar at the bottom.
    *   The selected image will appear in the **Image Preview** box.
    *   The extracted data will be displayed in the **Extraction Results** box.
    *   All extracted data is also automatically saved to an Excel file named `extracted_data.xlsx` in the application's folder.

## 3. Database Management Tab

This tab allows you to manage a local database that maps Manufacturer Part Numbers (MPNs) to your company-specific Customer Part Numbers (CPNs). When the application extracts an MPN, it will automatically look up the corresponding CPN in this database and add it to the results.

### Features:

-   **Open Full Database Manager:**
    *   Click this button to open a new window where you can view, add, edit, and delete individual MPN-CPN records.
    *   **To Add:** Type the MPN and CPN in the fields and click "Add".
    *   **To Update:** Select a record from the list, modify the MPN or CPN in the fields below, and click "Update".
    *   **To Delete:** Select a record from the list and click "Delete".

-   **Import from Excel:**
    *   Click this to import a list of MPN-CPN mappings from an Excel file.
    *   The Excel file **must** have two columns named `mpn` and `cpn`.
    *   The import process will skip any MPNs that already exist in the database.

-   **Export to Excel:**
    *   Click this to export all records from the database into a new Excel file. This is useful for backing up or sharing your MPN-CPN list.

## 4. Model Training Tab

This is the most advanced feature of the application. By providing images and labeling them with the correct data, you create a "ground truth" dataset. While this version of the application doesn't automatically retrain the OCR model, it saves your labeled data to the database. This data can be used in the future to fine-tune a custom OCR model for even better accuracy on your specific hardware components.

### How to Train:

1.  **Start a New Session:**
    *   Click the **Start New Training Session** button.
    *   A file dialog will open. You can select one or multiple images to label.

2.  **Label the Images:**
    *   The training interface will appear.
    *   The first image you selected will be shown in the **Image Preview**.
    *   On the right, you will see entry fields for all the data points (PN, MPN, CPN, etc.).
    *   Fill in the fields with the correct data as seen on the image. If a field is not present on the image, leave it blank.

3.  **Navigate and Save:**
    *   Use the **< Previous** and **Next >** buttons to navigate between the images you selected. Your labels for the current image are automatically saved when you navigate away.
    *   The status label shows you which image you are currently viewing (e.g., "Image 1/10").

4.  **Finish the Session:**
    *   When you are done labeling all the images, click the **Finish & Save Session** button.
    *   All the labels you provided will be saved to the `training_data` table in the application's database.
    *   The original image files will be copied to the `training_data` directory for future reference.

Thank you for using the Hardware Component OCR Data Extractor!
