#!/usr/bin/env python3
"""
OCR Hardware Component Data Extraction Application
Extracts PN, MPN, CPN, SSN, SN from hardware component labels
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import cv2
import numpy as np
from PIL import Image, ImageTk, ImageEnhance
import pytesseract
try:
    from pyzbar import pyzbar
    BARCODE_AVAILABLE = True
except ImportError:
    BARCODE_AVAILABLE = False
    print("Warning: pyzbar not available. Barcode detection will be disabled.")
import pandas as pd
import re
import os
import datetime
from pathlib import Path
import threading
import platform
import subprocess
import sqlite3

# Configure Tesseract path for Windows
def configure_tesseract_path():
    """Configure Tesseract executable path for different operating systems"""
    system = platform.system()
    
    if system == "Windows":
        # Common Windows installation paths
        possible_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(os.getenv('USERNAME', '')),
            r"C:\tesseract\tesseract.exe",
        ]
        
        # Try to find tesseract in common locations
        for path in possible_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                print(f"Tesseract found at: {path}")
                return True
        
        # If not found in common locations, try to find it
        try:
            result = subprocess.run(['where', 'tesseract'], capture_output=True, text=True)
            if result.returncode == 0:
                tesseract_path = result.stdout.strip().split('\n')[0]
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                print(f"Tesseract found at: {tesseract_path}")
                return True
        except:
            pass
        
        print("Warning: Tesseract not found. Please install Tesseract OCR or set the path manually.")
        print("If you have Tesseract installed, you can set the path in the code:")
        print("pytesseract.pytesseract.tesseract_cmd = r'C:\\path\\to\\tesseract.exe'")
        return False
    
    # For Linux/Mac, tesseract should be in PATH
    try:
        subprocess.run(['tesseract', '--version'], capture_output=True, check=True)
        print("Tesseract found in system PATH")
        return True
    except:
        print("Warning: Tesseract not found. Please install tesseract-ocr package.")
        return False

# Configure Tesseract on startup
configure_tesseract_path()

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hardware Component OCR Data Extractor")
        self.root.geometry("800x600")
        
        # Initialize variables
        self.current_image = None
        self.deep_scan_enabled = tk.BooleanVar(value=False)
        self.camera = None
        self.camera_active = False
        self.excel_file = "extracted_data.xlsx"
        self.training_data_dir = "training_data"
        self.db_file = "cpn_database.db"
        
        # Create training data directory if it doesn't exist
        Path(self.training_data_dir).mkdir(exist_ok=True)
        
        # Setup GUI
        self.setup_gui()
        
        # Initialize Database
        self.init_database()
        
    def setup_gui(self):
        """Setup the main GUI interface with tabs."""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Hardware Component OCR Data Extractor",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 10), sticky="ew")

        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky="nsew")

        # Create tabs
        self.tab_ocr = ttk.Frame(self.notebook, padding="10")
        self.tab_db = ttk.Frame(self.notebook, padding="10")
        self.tab_training = ttk.Frame(self.notebook, padding="10")

        self.notebook.add(self.tab_ocr, text="OCR Extraction")
        self.notebook.add(self.tab_db, text="Database Management")
        self.notebook.add(self.tab_training, text="Model Training")

        # --- Populate OCR Tab ---
        self.setup_ocr_tab()

        # --- Populate DB Tab ---
        self.setup_db_tab()

        # --- Populate Training Tab ---
        self.setup_training_tab()

        # Status bar
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        status_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky="ew")
        
        self.status_label = ttk.Label(status_frame, text="Ready")
        self.status_label.grid(row=1, column=0, sticky="ew")

    def setup_ocr_tab(self):
        """Setup the content of the OCR Extraction tab."""
        self.tab_ocr.columnconfigure(1, weight=1)
        self.tab_ocr.rowconfigure(1, weight=1)

        # Top frame for buttons
        controls_frame = ttk.Frame(self.tab_ocr)
        controls_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10))

        btn_upload = ttk.Button(controls_frame, text="Upload Picture & Extract", command=self.upload_image)
        btn_upload.pack(side="left", padx=(0, 10))

        btn_camera = ttk.Button(controls_frame, text="Camera Capture & Extract", command=self.camera_capture)
        btn_camera.pack(side="left")

        deep_scan_cb = ttk.Checkbutton(controls_frame, text="Enable Deep Scan (Slower)", variable=self.deep_scan_enabled)
        deep_scan_cb.pack(side="left", padx=(20, 0))

        # Image and Results frames
        image_frame = ttk.LabelFrame(self.tab_ocr, text="Image Preview", padding="5")
        image_frame.grid(row=1, column=0, pady=5, sticky="nsew")
        
        self.image_label = ttk.Label(image_frame, text="No image loaded")
        self.image_label.pack()

        results_frame = ttk.LabelFrame(self.tab_ocr, text="Extraction Results", padding="5")
        results_frame.grid(row=1, column=1, pady=5, sticky="nsew")
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

        self.results_text = scrolledtext.ScrolledText(results_frame, height=15, width=60)
        self.results_text.grid(row=0, column=0, sticky="nsew")

    def setup_db_tab(self):
        """Setup the content of the Database Management tab."""
        db_controls_frame = ttk.LabelFrame(self.tab_db, text="MPN-CPN Database", padding="10")
        db_controls_frame.pack(fill="x", pady=5)

        btn_db_manage = ttk.Button(db_controls_frame, text="Open Full Database Manager", command=self.open_db_management_window)
        btn_db_manage.pack(pady=5)
        
        io_frame = ttk.Frame(db_controls_frame)
        io_frame.pack(pady=10)

        import_btn = ttk.Button(io_frame, text="Import from Excel", command=self.import_from_excel)
        import_btn.pack(side="left", padx=5)

        export_btn = ttk.Button(io_frame, text="Export to Excel", command=self.export_to_excel)
        export_btn.pack(side="left", padx=5)

    def setup_training_tab(self):
        """Setup the content of the Model Training tab."""
        self.tab_training.columnconfigure(1, weight=1)
        self.tab_training.rowconfigure(1, weight=1)

        # Top control frame
        train_controls_frame = ttk.Frame(self.tab_training)
        train_controls_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10))
        btn_start_training = ttk.Button(train_controls_frame, text="Start New Training Session", command=self.train_model)
        btn_start_training.pack(side="left")

        # Main content frame, initially hidden
        self.training_content_frame = ttk.Frame(self.tab_training)
        self.training_content_frame.grid(row=1, column=0, columnspan=2, sticky='nsew')
        self.training_content_frame.grid_remove() # Hide it initially

        self.training_content_frame.columnconfigure(1, weight=1)
        self.training_content_frame.rowconfigure(0, weight=1)

        # Image display frame
        img_frame = ttk.LabelFrame(self.training_content_frame, text="Image Preview", padding="5")
        img_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.training_image_label = ttk.Label(img_frame, text="Select images to begin training.")
        self.training_image_label.pack()
        
        # Labeling frame
        label_frame = ttk.LabelFrame(self.training_content_frame, text="Data Labels", padding="10")
        label_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        self.training_entries = {}
        fields = ['PN', 'Part_Number', 'MPN', 'CPN', 'SSN', 'SN']
        for i, field in enumerate(fields):
            row = i // 2
            col = (i % 2) * 2
            ttk.Label(label_frame, text=f"{field}:").grid(row=row, column=col, sticky="w", padx=5, pady=2)
            entry = ttk.Entry(label_frame, width=30)
            entry.grid(row=row, column=col+1, padx=5, pady=2)
            self.training_entries[field] = entry
        
        # Navigation frame
        nav_frame = ttk.Frame(self.training_content_frame)
        nav_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)
        
        self.prev_btn = ttk.Button(nav_frame, text="< Previous", command=self.prev_training_image, state="disabled")
        self.prev_btn.pack(side=tk.LEFT, padx=5)
        
        self.next_btn = ttk.Button(nav_frame, text="Next >", command=self.next_training_image, state="disabled")
        self.next_btn.pack(side=tk.LEFT, padx=5)
        
        self.training_status = ttk.Label(nav_frame, text="")
        self.training_status.pack(side=tk.LEFT, padx=20)

        self.finish_btn = ttk.Button(nav_frame, text="Finish & Save Session", command=self.finish_training_session, state="disabled")
        self.finish_btn.pack(side=tk.RIGHT, padx=5)

    def init_database(self):
        """Initialize the SQLite database and create tables if they don't exist."""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            # Create mpn_cpn_map table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS mpn_cpn_map (
                    mpn TEXT PRIMARY KEY,
                    cpn TEXT NOT NULL
                )
            ''')
            # Create training_data table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS training_data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    image_path TEXT NOT NULL UNIQUE,
                    pn TEXT,
                    part_number TEXT,
                    mpn TEXT,
                    cpn TEXT,
                    ssn TEXT,
                    sn TEXT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            conn.commit()
            conn.close()
            print(f"Database '{self.db_file}' initialized successfully.")
        except Exception as e:
            print(f"Error initializing database: {str(e)}")
            
    def open_db_management_window(self):
        DBManagementWindow(self.root, self.db_file)

    def export_to_excel(self):
        """Export the database contents to an Excel file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export Database to Excel"
        )
        if not file_path:
            return
        try:
            conn = sqlite3.connect(self.db_file)
            df = pd.read_sql_query("SELECT * FROM mpn_cpn_map", conn)
            conn.close()
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Export Successful", f"Database successfully exported to\n{file_path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export database: {e}", parent=self.root)

    def import_from_excel(self):
        """Import records from an Excel file into the database."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Import from Excel"
        )
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path)
            if 'mpn' not in df.columns or 'cpn' not in df.columns:
                messagebox.showerror("Import Error", "Excel file must contain 'mpn' and 'cpn' columns.", parent=self.root)
                return

            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            imported_count = 0
            skipped_count = 0
            for index, row in df.iterrows():
                mpn = str(row['mpn'])
                cpn = str(row['cpn'])
                try:
                    # Use INSERT OR IGNORE to skip duplicates based on the PRIMARY KEY (mpn)
                    cursor.execute("INSERT OR IGNORE INTO mpn_cpn_map (mpn, cpn) VALUES (?, ?)", (mpn, cpn))
                    if cursor.rowcount > 0:
                        imported_count += 1
                    else:
                        skipped_count += 1
                except sqlite3.Error as insert_error:
                    print(f"Skipping row due to DB error: {insert_error}")
                    skipped_count += 1
            conn.commit()
            conn.close()
            summary_message = f"Import Complete!\n\nSuccessfully imported: {imported_count} records.\nSkipped (already exist or error): {skipped_count} records."
            messagebox.showinfo("Import Summary", summary_message, parent=self.root)
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import from Excel file: {e}", parent=self.root)

    def upload_image(self):
        """Handle image upload and processing"""
        file_path = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        
        if file_path:
            self.process_image(file_path)
            
    def camera_capture(self):
        """Handle camera capture"""
        if not self.camera_active:
            self.start_camera()
        else:
            self.stop_camera()
            
    def start_camera(self):
        """Start camera capture"""
        try:
            self.camera = cv2.VideoCapture(0)
            if not self.camera.isOpened():
                messagebox.showerror("Error", "Could not open camera")
                return
                
            self.camera_active = True
            self.update_status("Camera active - Click to capture image")
            
            # Create camera window
            self.camera_window = tk.Toplevel(self.root)
            self.camera_window.title("Camera Capture")
            self.camera_window.geometry("640x480")
            
            self.camera_label = ttk.Label(self.camera_window)
            self.camera_label.pack()
            
            capture_btn = ttk.Button(self.camera_window, text="Capture Image", 
                                   command=self.capture_image)
            capture_btn.pack(pady=10)
            
            close_btn = ttk.Button(self.camera_window, text="Close Camera", 
                                 command=self.stop_camera)
            close_btn.pack(pady=5)
            
            self.camera_window.protocol("WM_DELETE_WINDOW", self.stop_camera)
            
            self.update_camera_feed()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start camera: {str(e)}")
            
    def update_camera_feed(self):
        """Update camera feed in real-time"""
        if self.camera_active and self.camera is not None:
            ret, frame = self.camera.read()
            if ret:
                # Convert to RGB and resize for display
                frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                frame_resized = cv2.resize(frame_rgb, (640, 480))
                
                # Convert to PIL Image and display
                image = Image.fromarray(frame_resized)
                photo = ImageTk.PhotoImage(image)
                
                if hasattr(self, 'camera_label'):
                    self.camera_label.configure(image=photo)
                    self.camera_label.photo = photo  # Keep a reference
                    
            self.root.after(50, self.update_camera_feed)
            
    def capture_image(self):
        """Capture current frame from camera"""
        if self.camera is not None:
            ret, frame = self.camera.read()
            if ret:
                # Save captured image
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"captured_{timestamp}.jpg"
                cv2.imwrite(filename, frame)
                
                self.stop_camera()
                self.process_image(filename)
                
    def stop_camera(self):
        """Stop camera capture"""
        self.camera_active = False
        if self.camera is not None:
            self.camera.release()
            self.camera = None
            
        if hasattr(self, 'camera_window'):
            self.camera_window.destroy()
            
        self.update_status("Camera stopped")
        
    def train_model(self):
        """Handle training data upload with labeling."""
        files = filedialog.askopenfilenames(
            title="Select Training Images",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        if not files:
            return

        self.training_files = files
        self.current_training_index = 0
        self.training_labels = [{} for _ in files] # Initialize labels list

        # Show the training UI
        self.training_content_frame.grid()
        self.finish_btn.config(state="normal")
        
        self.load_training_image()

    def load_training_image(self):
        """Load current training image for labeling."""
        if not hasattr(self, 'training_files') or not (0 <= self.current_training_index < len(self.training_files)):
            return

        file_path = self.training_files[self.current_training_index]
        try:
            image = Image.open(file_path)
            image.thumbnail((400, 300), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            self.training_image_label.configure(image=photo, text="")
            self.training_image_label.photo = photo
        except Exception as e:
            self.training_image_label.configure(image=None, text=f"Error loading image:\n{e}")

        total = len(self.training_files)
        current = self.current_training_index + 1
        filename = os.path.basename(file_path)
        self.training_status.configure(text=f"Image {current}/{total}: {filename}")

        # Load existing labels for this image
        labels = self.training_labels[self.current_training_index]
        for field, entry in self.training_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, labels.get(field, ""))
            
        # Update button states
        self.prev_btn.configure(state="normal" if self.current_training_index > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.current_training_index < len(self.training_files) - 1 else "disabled")

    def prev_training_image(self):
        """Go to previous training image"""
        if self.current_training_index > 0:
            self.save_current_labels()
            self.current_training_index -= 1
            self.load_training_image()

    def next_training_image(self):
        """Go to next training image"""
        if self.current_training_index < len(self.training_files) - 1:
            self.save_current_labels()
            self.current_training_index += 1
            self.load_training_image()

    def save_current_labels(self):
        """Save labels for current image to the instance variable."""
        if not hasattr(self, 'training_files'):
            return
        labels = {}
        for field, entry in self.training_entries.items():
            labels[field] = entry.get().strip()
        self.training_labels[self.current_training_index] = labels

    def finish_training_session(self):
        """Save all collected labels from the session to the database."""
        self.save_current_labels() # Save the very last image's labels
        
        import shutil
        saved_count, skipped_count = 0, 0
        
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            for i, file_path in enumerate(self.training_files):
                labels = self.training_labels[i]
                if not any(labels.values()):
                    skipped_count += 1
                    continue

                filename = os.path.basename(file_path)
                dest_path = os.path.join(self.training_data_dir, filename)

                # Copy file to training dir, overwriting if exists
                shutil.copy2(file_path, dest_path)
                db_path = Path(dest_path).as_posix() # Use posix paths for db consistency

                try:
                    cursor.execute('''
                        INSERT OR REPLACE INTO training_data
                        (image_path, pn, part_number, mpn, cpn, ssn, sn, timestamp)
                        VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    ''', (db_path, labels.get('PN',''), labels.get('Part_Number',''), labels.get('MPN',''), labels.get('CPN',''), labels.get('SSN',''), labels.get('SN','')))
                    saved_count += 1
                except sqlite3.Error as e:
                    print(f"DB Error for {filename}: {e}")
                    skipped_count += 1

            conn.commit()
            conn.close()
            
            messagebox.showinfo("Training Session Saved", f"Saved {saved_count} labeled images to the database.\nSkipped {skipped_count} unlabeled images.")
            
        except Exception as e:
            messagebox.showerror("Training Save Error", f"Failed to save training session: {e}")
        
        # Hide the training UI and reset state
        self.training_content_frame.grid_remove()
        self.training_image_label.configure(image=None, text="Select images to begin training.")
        self.training_status.config(text="")
        self.prev_btn.config(state="disabled")
        self.next_btn.config(state="disabled")
        self.finish_btn.config(state="disabled")
        del self.training_files
        del self.training_labels
                      
    def preprocess_image(self, image_path):
        """Preprocess image for better OCR results with rotation correction"""
        image = cv2.imread(image_path)
        if image is None:
            raise ValueError("Could not load image")
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        processed_images = []

        if self.deep_scan_enabled.get():
            rotations = [0, 90, 180, 270]
            for angle in rotations:
                if angle == 0: rotated_gray = gray
                elif angle == 90: rotated_gray = cv2.rotate(gray, cv2.ROTATE_90_CLOCKWISE)
                elif angle == 180: rotated_gray = cv2.rotate(gray, cv2.ROTATE_180)
                else: rotated_gray = cv2.rotate(gray, cv2.ROTATE_90_COUNTERCLOCKWISE)

                processed_images.append((f"rot_{angle}", rotated_gray))
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                enhanced = clahe.apply(rotated_gray)
                processed_images.append((f"rot_{angle}_enhanced", enhanced))
        else:
            processed_images.append(("original", gray))
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            enhanced = clahe.apply(gray)
            processed_images.append(("enhanced", enhanced))

        return processed_images, image
        
    def extract_text_ocr(self, processed_images):
        """Extract text using OCR from preprocessed images"""
        all_text = ""

        for name, img in processed_images:
            try:
                # Try different OCR configurations
                configs = [
                    '--psm 6',  # Uniform block of text
                    '--psm 8',  # Single word
                    '--psm 7',  # Single text line
                    '--psm 13', # Raw line
                ]
                
                for config in configs:
                    text = pytesseract.image_to_string(img, config=config)
                    if text.strip():
                        all_text += f"\n--- {name} ({config}) ---\n{text}\n"
                        
            except Exception as e:
                print(f"OCR error for {name}: {str(e)}")
                
        return all_text
        
    def detect_barcodes(self, image):
        """Detect and decode barcodes"""
        if not BARCODE_AVAILABLE:
            return ["Barcode detection unavailable (pyzbar not installed)"]
            
        try:
            barcodes = pyzbar.decode(image)
            barcode_data = []
            
            for barcode in barcodes:
                data = barcode.data.decode('utf-8')
                barcode_type = barcode.type
                barcode_data.append(f"{barcode_type}: {data}")
                
            return barcode_data
        except Exception as e:
            return [f"Barcode detection error: {str(e)}"]
            
        return barcode_data
        
    def extract_specific_data(self, text, barcode_data):
        """Extract specific hardware component data"""
        results = {
            'PN': [],
            'Part_Number': [],
            'MPN': [],
            'CPN': [],
            'SSN': [],
            'SN': [],
            'Cisco_Data': [],
            'Barcode_Data': barcode_data
        }
        
        # Patterns for different fields
        patterns = {
            'PN': [r'PN[:\s]+([A-Z0-9\-\.]+)', r'P/N[:\s]+([A-Z0-9\-\.]+)', r'Part\s*No[:\s]+([A-Z0-9\-\.]+)'],
            'Part_Number': [r'PART\s*NUMBER[:\s]+([A-Z0-9\-\.]+)', r'Part\s*Number[:\s]+([A-Z0-9\-\.]+)'],
            'MPN': [r'MPN[:\s]+([A-Z0-9\-\.]+)', r'Mfg\s*Part[:\s]+([A-Z0-9\-\.]+)'],
            'CPN': [r'CPN[:\s]+([A-Z0-9\-\.]+)', r'Customer\s*Part[:\s]+([A-Z0-9\-\.]+)'],
            'SSN': [r'SSN[:\s]+([A-Z0-9\-\.]+)'],
            'SN': [r'SN[:\s]+([A-Z0-9\-\.]+)', r'Serial[:\s]+([A-Z0-9\-\.]+)', r'S/N[:\s]+([A-Z0-9\-\.]+)']
        }
        
        # Extract data using patterns
        for field, field_patterns in patterns.items():
            for pattern in field_patterns:
                matches = re.findall(pattern, text, re.IGNORECASE)
                results[field].extend(matches)
        
        # Enhanced CPN detection - look for 16-XXXXXX-XX patterns in text
        cpn_patterns = [
            r'16-\d{6}-\d{2}',  # 16-101913-01 format
            r'16-[A-Z0-9]{6}-[A-Z0-9]{2}',  # 16-ABCDEF-01 format
            r'\b16-[A-Z0-9\-]{8,}\b'  # General 16- prefix patterns
        ]
        
        for pattern in cpn_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                if match not in results['CPN']:
                    results['CPN'].append(match)
        
        # Also check barcode data for CPN patterns
        for barcode in barcode_data:
            # Extract the actual barcode value
            if ': ' in barcode:
                barcode_value = barcode.split(': ')[-1]
            else:
                barcode_value = barcode
            
            # Check if barcode matches CPN patterns
            for pattern in cpn_patterns:
                if re.match(pattern, barcode_value, re.IGNORECASE):
                    if barcode_value not in results['CPN']:
                        results['CPN'].append(barcode_value)
                        
        # Look for Cisco logo and extract associated data
        cisco_indicators = ['cisco', 'CISCO', 'Cisco']
        for indicator in cisco_indicators:
            if indicator in text:
                # Extract text around Cisco mentions
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if indicator in line:
                        # Get surrounding lines
                        start = max(0, i-2)
                        end = min(len(lines), i+3)
                        cisco_context = '\n'.join(lines[start:end])
                        results['Cisco_Data'].append(cisco_context)
                        
        return results

    def lookup_cpn_from_mpn(self, mpn):
        """Look up CPN from MPN in the SQLite database."""
        if not os.path.exists(self.db_file):
            return None

        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("SELECT cpn FROM mpn_cpn_map WHERE mpn = ?", (mpn,))
            result = cursor.fetchone()
            conn.close()

            if result:
                return result[0]

        except Exception as e:
            print(f"Error looking up CPN in database: {str(e)}")

        return None

    def save_to_excel(self, image_file, extracted_data):
        """Save extracted data to Excel file"""
        try:
            # Read existing data
            try:
                df = pd.read_excel(self.excel_file)
                # Remove legacy All_Text column if it exists
                if 'All_Text' in df.columns:
                    df = df.drop(columns=['All_Text'])
            except FileNotFoundError:
                df = pd.DataFrame()
                
            # Prepare new row
            new_row = {
                'Timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'Image_File': os.path.basename(image_file),
                'PN': '; '.join(extracted_data['PN']) if extracted_data['PN'] else '',
                'Part_Number': '; '.join(extracted_data['Part_Number']) if extracted_data['Part_Number'] else '',
                'MPN': '; '.join(extracted_data['MPN']) if extracted_data['MPN'] else '',
                'CPN': '; '.join(extracted_data['CPN']) if extracted_data['CPN'] else '',
                'SSN': '; '.join(extracted_data['SSN']) if extracted_data['SSN'] else '',
                'SN': '; '.join(extracted_data['SN']) if extracted_data['SN'] else '',
                'Cisco_Data': '; '.join(extracted_data['Cisco_Data']) if extracted_data['Cisco_Data'] else '',
                'Barcode_Data': '; '.join(extracted_data['Barcode_Data']) if extracted_data['Barcode_Data'] else ''
            }
            
            # Append to dataframe
            new_df = pd.DataFrame([new_row])
            df = pd.concat([df, new_df], ignore_index=True)
            
            # Save to Excel
            df.to_excel(self.excel_file, index=False)
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save to Excel: {str(e)}")
            return False
            
    def display_image(self, image_path):
        """Display image in the GUI"""
        try:
            # Load and resize image for display
            image = Image.open(image_path)

            # Calculate size to fit in display area
            display_width = 500
            display_height = 300
            image.thumbnail((display_width, display_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # Update label
            self.image_label.configure(image=photo, text="")
            self.image_label.photo = photo  # Keep a reference
            
        except Exception as e:
            self.image_label.configure(text=f"Error loading image: {str(e)}")
            
    def update_results(self, extracted_data, all_text, barcode_data):
        """Update results display with clear formatting"""
        self.results_text.delete(1.0, tk.END)
        
        results = "📋 === EXTRACTION RESULTS ===\n\n"
        
        # Display extracted fields with better formatting
        field_names = {
            'PN': '🔧 Part Number (PN)',
            'Part_Number': '📦 Part Number',
            'MPN': '🏭 Manufacturer Part Number (MPN)',
            'CPN': '👤 Customer Part Number (CPN)',
            'SSN': '🔢 SSN',
            'SN': '📋 Serial Number (SN)',
            'Cisco_Data': '🌐 Cisco Data'
        }
        
        found_data = False
        for field, values in extracted_data.items():
            if field != 'Barcode_Data' and values:
                field_display = field_names.get(field, field)
                results += f"{field_display}:\n"
                for value in values:
                    results += f"   ✓ {value}\n"
                results += "\n"
                found_data = True
                
        # Display barcode data
        if barcode_data and any(barcode_data):
            results += f"📱 Barcode Data:\n"
            for barcode in barcode_data:
                if barcode and not barcode.startswith("Barcode detection"):
                    results += f"   ✓ {barcode}\n"
            results += "\n"
            found_data = True
        
        if not found_data:
            results += "⚠️ No component data found in this image.\n"
            results += "Try adjusting image orientation or quality.\n\n"
        else:
            results += f"✅ Data extraction completed successfully!\n"
            results += f"💾 Results saved to {self.excel_file}\n\n"
        
        # Show sample of extracted text for debugging
        if all_text and len(all_text.strip()) > 0:
            sample_text = all_text[:300].replace('\n', ' ').strip()
            if sample_text:
                results += f"📝 Sample Extracted Text:\n{sample_text}..."
        
        self.results_text.insert(1.0, results)
        
    def update_status(self, message):
        """Update status label"""
        self.status_label.configure(text=message)
        self.root.update()
        
    def process_image(self, image_path):
        """Main image processing function"""
        def process():
            try:
                # All heavy processing in background thread
                self.root.after(0, lambda: self.progress.start())
                self.root.after(0, lambda: self.update_status("Processing image..."))
                
                # Display image
                self.root.after(0, lambda: self.display_image(image_path))

                # Preprocess image
                self.root.after(0, lambda: self.update_status("Preprocessing image..."))
                processed_images, original_image = self.preprocess_image(image_path)
                
                # Extract text using OCR
                self.root.after(0, lambda: self.update_status("Extracting text..."))
                all_text = self.extract_text_ocr(processed_images)
                
                # Detect barcodes
                self.root.after(0, lambda: self.update_status("Detecting barcodes..."))
                barcode_data = self.detect_barcodes(original_image)
                
                # Extract specific data
                self.root.after(0, lambda: self.update_status("Extracting component data..."))
                extracted_data = self.extract_specific_data(all_text, barcode_data)
                
                # Perform MPN to CPN lookup
                if extracted_data.get('MPN'):
                    self.root.after(0, lambda: self.update_status("Looking up CPN from MPN..."))
                    found_cpns = []
                    for mpn in extracted_data['MPN']:
                        cpn = self.lookup_cpn_from_mpn(mpn)
                        if cpn:
                            found_cpns.append(cpn)

                    if found_cpns:
                        # Override any existing CPNs with the looked-up values
                        extracted_data['CPN'] = found_cpns

                # Save to Excel
                self.root.after(0, lambda: self.update_status("Saving to Excel..."))
                success = self.save_to_excel(image_path, extracted_data)
                
                # Update UI in main thread
                def update_ui():
                    if success:
                        self.update_status(f"Data saved to {self.excel_file}")
                    self.update_results(extracted_data, all_text, barcode_data)
                    self.progress.stop()
                    self.update_status("Processing completed successfully")
                    
                self.root.after(0, update_ui)
                
            except Exception as e:
                # Handle errors in main thread
                def handle_error():
                    self.progress.stop()
                    self.update_status("Processing failed")
                    messagebox.showerror("Error", f"Processing failed: {str(e)}")
                    
                self.root.after(0, handle_error)
                
        # Run in separate thread to prevent GUI freezing
        thread = threading.Thread(target=process)
        thread.daemon = True
        thread.start()

# --- DB Management Window ---
class DBManagementWindow(tk.Toplevel):
    def __init__(self, parent, db_file):
        super().__init__(parent)
        self.db_file = db_file
        self.title("Database Management")
        self.geometry("800x600")

        self.create_widgets()
        self.load_data()

    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Top frame for import/export buttons
        io_frame = ttk.Frame(main_frame)
        io_frame.pack(fill="x", pady=5)

        import_btn = ttk.Button(io_frame, text="Import from Excel", command=self.import_from_excel)
        import_btn.pack(side="left", padx=5)

        export_btn = ttk.Button(io_frame, text="Export to Excel", command=self.export_to_excel)
        export_btn.pack(side="left", padx=5)

        # Frame for the Treeview
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # Treeview to display data
        self.tree = ttk.Treeview(tree_frame, columns=("mpn", "cpn"), show="headings")
        self.tree.heading("mpn", text="MPN")
        self.tree.heading("cpn", text="CPN")
        self.tree.pack(side="left", fill="both", expand=True)

        # Scrollbar for the Treeview
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Frame for entry fields and buttons
        entry_frame = ttk.LabelFrame(main_frame, text="Manage Record")
        entry_frame.pack(pady=10, padx=10, fill="x")

        # Entry fields
        ttk.Label(entry_frame, text="MPN:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.mpn_entry = ttk.Entry(entry_frame, width=40)
        self.mpn_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(entry_frame, text="CPN:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cpn_entry = ttk.Entry(entry_frame, width=40)
        self.cpn_entry.grid(row=1, column=1, padx=5, pady=5)

        # Buttons for CRUD operations
        crud_frame = ttk.Frame(entry_frame)
        crud_frame.grid(row=0, column=2, rowspan=2, padx=20)

        add_btn = ttk.Button(crud_frame, text="Add", command=self.add_record)
        add_btn.pack(pady=2, fill="x")

        update_btn = ttk.Button(crud_frame, text="Update", command=self.update_record)
        update_btn.pack(pady=2, fill="x")

        delete_btn = ttk.Button(crud_frame, text="Delete", command=self.delete_record)
        delete_btn.pack(pady=2, fill="x")

        clear_btn = ttk.Button(crud_frame, text="Clear Fields", command=self.clear_fields)
        clear_btn.pack(pady=2, fill="x")

        # Bind tree selection to a method
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def load_data(self):
        # Clear existing data
        for row in self.tree.get_children():
            self.tree.delete(row)
        # Load new data
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM mpn_cpn_map ORDER BY mpn")
            for row in cursor.fetchall():
                self.tree.insert("", "end", values=row)
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data from database: {e}", parent=self)

    def add_record(self):
        mpn = self.mpn_entry.get().strip()
        cpn = self.cpn_entry.get().strip()
        if not mpn or not cpn:
            messagebox.showwarning("Input Error", "MPN and CPN fields cannot be empty.", parent=self)
            return
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO mpn_cpn_map (mpn, cpn) VALUES (?, ?)", (mpn, cpn))
            conn.commit()
            conn.close()
            self.load_data()
            self.clear_fields()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"MPN '{mpn}' already exists.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add record: {e}", parent=self)

    def update_record(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a record to update.", parent=self)
            return

        original_mpn = self.tree.item(selected_item, "values")[0]
        mpn = self.mpn_entry.get().strip()
        cpn = self.cpn_entry.get().strip()

        if not mpn or not cpn:
            messagebox.showwarning("Input Error", "MPN and CPN fields cannot be empty.", parent=self)
            return

        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("UPDATE mpn_cpn_map SET mpn = ?, cpn = ? WHERE mpn = ?", (mpn, cpn, original_mpn))
            conn.commit()
            conn.close()
            self.load_data()
            self.clear_fields()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update record: {e}", parent=self)

    def delete_record(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a record to delete.", parent=self)
            return

        mpn = self.tree.item(selected_item, "values")[0]
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete MPN '{mpn}'?", parent=self):
            try:
                conn = sqlite3.connect(self.db_file)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM mpn_cpn_map WHERE mpn = ?", (mpn,))
                conn.commit()
                conn.close()
                self.load_data()
                self.clear_fields()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete record: {e}", parent=self)

    def on_select(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            mpn, cpn = self.tree.item(selected_item, "values")
            self.clear_fields()
            self.mpn_entry.insert(0, mpn)
            self.cpn_entry.insert(0, cpn)

    def clear_fields(self):
        self.mpn_entry.delete(0, tk.END)
        self.cpn_entry.delete(0, tk.END)

    def export_to_excel(self):
        """Export the database contents to an Excel file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export Database to Excel"
        )

        if not file_path:
            return # User cancelled

        try:
            conn = sqlite3.connect(self.db_file)
            df = pd.read_sql_query("SELECT * FROM mpn_cpn_map", conn)
            conn.close()

            df.to_excel(file_path, index=False)
            messagebox.showinfo("Export Successful", f"Database successfully exported to\n{file_path}", parent=self)

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export database: {e}", parent=self)

    def import_from_excel(self):
        """Import records from an Excel file into the database."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Import from Excel"
        )

        if not file_path:
            return # User cancelled

        try:
            df = pd.read_excel(file_path)
            if 'mpn' not in df.columns or 'cpn' not in df.columns:
                messagebox.showerror("Import Error", "Excel file must contain 'mpn' and 'cpn' columns.", parent=self)
                return

            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            imported_count = 0
            skipped_count = 0

            for index, row in df.iterrows():
                mpn = str(row['mpn'])
                cpn = str(row['cpn'])
                try:
                    cursor.execute("INSERT OR IGNORE INTO mpn_cpn_map (mpn, cpn) VALUES (?, ?)", (mpn, cpn))
                    if cursor.rowcount > 0:
                        imported_count += 1
                    else:
                        skipped_count += 1
                except sqlite3.Error:
                    skipped_count += 1

            conn.commit()
            conn.close()

            self.load_data() # Refresh the view

            summary_message = f"Import Complete!\n\n"
            summary_message += f"Successfully imported: {imported_count} records.\n"
            summary_message += f"Skipped (already exist): {skipped_count} records."
            messagebox.showinfo("Import Summary", summary_message, parent=self)

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import from Excel file: {e}", parent=self)

def main():
    """Main application entry point"""
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()