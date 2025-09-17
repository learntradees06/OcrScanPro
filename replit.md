# Hardware Component OCR Data Extractor

## Overview

This project is a complete desktop application built with Python that extracts component identification data (Part Numbers, Manufacturer Part Numbers, Component Part Numbers, Serial Numbers, etc.) from hardware component labels using Optical Character Recognition (OCR). The application features a comprehensive Tkinter-based GUI with three main functions: image upload processing, live camera capture, and training data management. All extracted data is automatically saved to Excel files with timestamps for continuous data collection and analysis.

## Current Status

✅ **COMPLETED** - Fully functional OCR application with all requested features:
- Local Python desktop application (no web dependencies)
- Three main functions: Upload Image, Camera Capture, Train Model
- Advanced OCR with rotation correction (0°, 90°, 180°, 270°)
- Intelligent data extraction for PN, MPN, CPN, SSN, SN fields
- Cisco logo detection and associated data extraction
- Excel export with continuous data append functionality
- Complete training module with labeling interface
- Thread-safe GUI operations

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### GUI Framework
- **Tkinter-based desktop application**: Chosen for its simplicity and built-in availability with Python installations
- **Main window with image preview**: Allows users to visualize the image being processed before extraction
- **Tabbed interface design**: Separates different functionalities like file processing, camera input, and settings

### Image Processing Pipeline
- **OpenCV for image preprocessing**: Handles grayscale conversion, contrast enhancement, noise reduction, and thresholding
- **Multiple preprocessing strategies**: Applies various image enhancement techniques (CLAHE, Gaussian blur, OTSU thresholding) to improve OCR accuracy
- **PIL/Pillow integration**: Manages image display and format conversions for the GUI

### OCR Engine
- **Tesseract OCR**: Primary text extraction engine with configurable page segmentation modes
- **Multiple OCR configurations**: Tests different PSM (Page Segmentation Mode) settings to maximize text extraction accuracy
- **Barcode detection**: Optional pyzbar integration for reading barcodes and QR codes when available

### Data Extraction Logic
- **Regex pattern matching**: Identifies component identifiers (PN, MPN, CPN, SSN, SN) using predefined patterns
- **Text parsing algorithms**: Processes raw OCR output to extract structured component data
- **Training data collection**: Stores processed images and results for potential machine learning improvements

### Data Storage
- **Excel file output**: Uses pandas to write extracted data to Excel format
- **Structured data schema**: Organizes extracted information with timestamps, image sources, and component identifiers
- **Incremental data appending**: Adds new extractions to existing Excel files without overwriting

### Camera Integration
- **OpenCV camera interface**: Enables live camera feed for real-time component scanning
- **Threading implementation**: Prevents GUI blocking during camera operations and image processing

## External Dependencies

### Core Libraries
- **OpenCV (cv2)**: Image processing, camera interface, and computer vision operations
- **Tesseract OCR (pytesseract)**: Text extraction from images
- **PIL/Pillow**: Image manipulation and format handling
- **pandas**: Data structure management and Excel file operations
- **numpy**: Numerical operations and array handling

### Optional Dependencies
- **pyzbar**: Barcode and QR code detection (gracefully handles absence)

### System Requirements
- **Tesseract OCR binary**: Must be installed on the system and accessible via PATH
- **Camera hardware**: For live scanning functionality
- **Excel-compatible software**: For viewing extracted data results

### File System Dependencies
- **Training data directory**: Local storage for processed images and learning data
- **Excel output files**: Local file system for data persistence
- **Image file support**: Supports common image formats (PNG, JPG, BMP) through PIL