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
        self.image_path = None
        self.original_pil_image = None
        self.display_pil_image = None
        self.auto_rotate_enabled = tk.BooleanVar(value=False)
        self.camera = None
        self.camera_active = False
        self.excel_file = "extracted_data.xlsx"
        self.training_data_dir = "training_data"
        
        # Create training data directory if it doesn't exist
        Path(self.training_data_dir).mkdir(exist_ok=True)
        
        # Setup GUI
        self.setup_gui()
        
        # Initialize Excel file
        self.init_excel_file()
        
    def setup_gui(self):
        """Setup the main GUI interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Title
        title_label = ttk.Label(main_frame, text="Hardware Component OCR Data Extractor", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Main action buttons
        btn_upload = ttk.Button(main_frame, text="1. Upload Picture",
                               command=self.upload_image, width=20)
        btn_upload.grid(row=1, column=0, padx=10, pady=10)
        
        btn_camera = ttk.Button(main_frame, text="2. Camera Capture",
                               command=self.camera_capture, width=20)
        btn_camera.grid(row=1, column=1, padx=10, pady=10)

        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#ccc")
        style.configure("Accent.TButton", font=("Arial", 12, "bold"), foreground="white", background="#007bff")

        self.extract_btn = ttk.Button(main_frame, text="3. Extract Data",
                                     style="Accent.TButton",
                                     command=self.extract_data_from_current_image,
                                     width=20, state="disabled")
        self.extract_btn.grid(row=1, column=2, padx=10, pady=10)

        auto_rotate_cb = ttk.Checkbutton(main_frame, text="Auto-rotate image (Slower)",
                                         variable=self.auto_rotate_enabled)
        auto_rotate_cb.grid(row=2, column=2, pady=5, sticky="w")
        
        # Image display frame
        image_frame = ttk.LabelFrame(main_frame, text="Image Preview", padding="5")
        image_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky="ew")
        
        self.image_label = ttk.Label(image_frame, text="No image loaded")
        self.image_label.pack(pady=5)

        # Rotation buttons frame
        rotation_frame = ttk.Frame(image_frame)
        rotation_frame.pack(pady=5)

        self.rotate_cw_btn = ttk.Button(rotation_frame, text="Rotate 90° CW", command=self.rotate_image_cw, state="disabled")
        self.rotate_cw_btn.pack(side=tk.LEFT, padx=5)

        self.rotate_ccw_btn = ttk.Button(rotation_frame, text="Rotate 90° CCW", command=self.rotate_image_ccw, state="disabled")
        self.rotate_ccw_btn.pack(side=tk.LEFT, padx=5)

        self.reset_btn = ttk.Button(rotation_frame, text="Reset", command=self.reset_image, state="disabled")
        self.reset_btn.pack(side=tk.LEFT, padx=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Extraction Results", padding="5")
        results_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky="nsew")
        
        self.results_text = scrolledtext.ScrolledText(results_frame, height=15, width=80)
        self.results_text.grid(row=0, column=0, sticky="nsew")
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
    def init_excel_file(self):
        """Initialize Excel file with headers if it doesn't exist"""
        if not os.path.exists(self.excel_file):
            columns = ['Timestamp', 'Image_File', 'PN', 'Part_Number', 'MPN', 'CPN', 'SSN', 'SN', 'Cisco_Data', 'Barcode_Data']
            df = pd.DataFrame(columns=columns)
            df.to_excel(self.excel_file, index=False)

    def rotate_image_cw(self):
        """Rotate the display image 90 degrees clockwise"""
        if self.display_pil_image:
            self.display_pil_image = self.display_pil_image.rotate(-90, expand=True)
            self.display_pil_image_in_gui(self.display_pil_image)

    def rotate_image_ccw(self):
        """Rotate the display image 90 degrees counter-clockwise"""
        if self.display_pil_image:
            self.display_pil_image = self.display_pil_image.rotate(90, expand=True)
            self.display_pil_image_in_gui(self.display_pil_image)

    def reset_image(self):
        """Reset the display image to the original"""
        if self.original_pil_image:
            self.display_pil_image = self.original_pil_image.copy()
            self.display_pil_image_in_gui(self.display_pil_image)

    def update_rotation_buttons_state(self, state):
        """Enable or disable the rotation buttons"""
        self.rotate_cw_btn.config(state=state)
        self.rotate_ccw_btn.config(state=state)
        self.reset_btn.config(state=state)
            
    def upload_image(self):
        """Handle image upload and processing"""
        file_path = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        
        if file_path:
            self.image_path = file_path
            self.original_pil_image = Image.open(file_path)
            self.display_pil_image = self.original_pil_image.copy()
            self.display_pil_image_in_gui(self.display_pil_image)
            self.update_rotation_buttons_state("normal")
            self.extract_btn.config(state="normal")
            
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
                
                # Load the captured image and enable extraction
                self.image_path = filename
                self.original_pil_image = Image.open(filename)
                self.display_pil_image = self.original_pil_image.copy()
                self.display_pil_image_in_gui(self.display_pil_image)
                self.update_rotation_buttons_state("normal")
                self.extract_btn.config(state="normal")

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
        """Handle training data upload with labeling"""
        files = filedialog.askopenfilenames(
            title="Select Training Images",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        
        if files:
            self.open_training_window(files)
            
    def open_training_window(self, files):
        """Open training window with labeling interface"""
        training_window = tk.Toplevel(self.root)
        training_window.title("Training Data Upload and Labeling")
        training_window.geometry("800x600")
        
        # Current file index
        self.current_training_index = 0
        self.training_files = files
        self.training_labels = []
        
        # Image display frame
        img_frame = ttk.LabelFrame(training_window, text="Image Preview", padding="5")
        img_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.training_image_label = ttk.Label(img_frame, text="Loading...")
        self.training_image_label.pack()
        
        # Labeling frame
        label_frame = ttk.LabelFrame(training_window, text="Data Labels", padding="10")
        label_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create label entry fields
        self.training_entries = {}
        fields = ['PN', 'Part_Number', 'MPN', 'CPN', 'SSN', 'SN']
        
        for i, field in enumerate(fields):
            row = i // 2
            col = (i % 2) * 2
            
            ttk.Label(label_frame, text=f"{field}:").grid(row=row, column=col, sticky="w", padx=5, pady=2)
            entry = ttk.Entry(label_frame, width=25)
            entry.grid(row=row, column=col+1, padx=5, pady=2)
            self.training_entries[field] = entry
        
        # Navigation frame
        nav_frame = ttk.Frame(training_window)
        nav_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.prev_btn = ttk.Button(nav_frame, text="< Previous", command=self.prev_training_image)
        self.prev_btn.pack(side=tk.LEFT, padx=5)
        
        self.next_btn = ttk.Button(nav_frame, text="Next >", command=self.next_training_image)
        self.next_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_labels_btn = ttk.Button(nav_frame, text="Save Labels", command=self.save_current_labels)
        self.save_labels_btn.pack(side=tk.LEFT, padx=20)
        
        self.finish_btn = ttk.Button(nav_frame, text="Finish Training", command=lambda: self.finish_training(training_window))
        self.finish_btn.pack(side=tk.RIGHT, padx=5)
        
        # Status label
        self.training_status = ttk.Label(training_window, text="")
        self.training_status.pack(pady=5)
        
        # Load first image
        self.load_training_image()
        
    def load_training_image(self):
        """Load current training image for labeling"""
        if 0 <= self.current_training_index < len(self.training_files):
            file_path = self.training_files[self.current_training_index]
            
            # Display image
            try:
                image = Image.open(file_path)
                image.thumbnail((400, 300), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(image)
                self.training_image_label.configure(image=photo, text="")
                self.training_image_label.photo = photo
            except Exception as e:
                self.training_image_label.configure(text=f"Error loading image: {str(e)}")
            
            # Update status
            total = len(self.training_files)
            current = self.current_training_index + 1
            filename = os.path.basename(file_path)
            self.training_status.configure(text=f"Image {current}/{total}: {filename}")
            
            # Load existing labels if available
            if self.current_training_index < len(self.training_labels):
                labels = self.training_labels[self.current_training_index]
                for field, entry in self.training_entries.items():
                    entry.delete(0, tk.END)
                    entry.insert(0, labels.get(field, ""))
            else:
                # Clear entries for new image
                for entry in self.training_entries.values():
                    entry.delete(0, tk.END)
                    
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
        """Save labels for current image"""
        labels = {}
        for field, entry in self.training_entries.items():
            labels[field] = entry.get().strip()
            
        # Extend labels list if needed
        while len(self.training_labels) <= self.current_training_index:
            self.training_labels.append({})
            
        self.training_labels[self.current_training_index] = labels
        
    def finish_training(self, window):
        """Finish training and save all data"""
        self.save_current_labels()
        
        # Save training data
        training_data = []
        import shutil
        
        for i, file_path in enumerate(self.training_files):
            filename = os.path.basename(file_path)
            dest_path = os.path.join(self.training_data_dir, filename)
            
            # Copy file
            shutil.copy2(file_path, dest_path)
            
            # Prepare training record
            labels = self.training_labels[i] if i < len(self.training_labels) else {}
            record = {
                'filename': filename,
                'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                **labels
            }
            training_data.append(record)
            
        # Save to CSV
        training_csv = os.path.join(self.training_data_dir, "training_labels.csv")
        df = pd.DataFrame(training_data)
        
        if os.path.exists(training_csv):
            # Append to existing file
            existing_df = pd.read_csv(training_csv)
            df = pd.concat([existing_df, df], ignore_index=True)
            
        df.to_csv(training_csv, index=False)
        
        messagebox.showinfo("Training Complete", 
                           f"Saved {len(self.training_files)} training images with labels to {self.training_data_dir}")
        
        window.destroy()
                      
    def preprocess_image(self, pil_image):
        """Preprocess a PIL image for better OCR results with advanced techniques."""
        # Convert PIL image to OpenCV format
        image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        processed_images = []

        # If auto-rotate is enabled, try all rotations
        if self.auto_rotate_enabled.get():
            rotations = [0, 90, 180, 270]
            for angle in rotations:
                if angle == 0:
                    rotated_gray = gray
                elif angle == 90:
                    rotated_gray = cv2.rotate(gray, cv2.ROTATE_90_CLOCKWISE)
                elif angle == 180:
                    rotated_gray = cv2.rotate(gray, cv2.ROTATE_180)
                elif angle == 270:
                    rotated_gray = cv2.rotate(gray, cv2.ROTATE_90_COUNTERCLOCKWISE)

                # Add the rotated grayscale image itself
                processed_images.append((f"rot_{angle}", rotated_gray))
                # Add an enhanced version
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                enhanced = clahe.apply(rotated_gray)
                processed_images.append((f"rot_{angle}_enhanced", enhanced))
        else:
            # If not auto-rotating, just use the advanced pipeline on the current orientation
            processed_images.append(("original", gray))
            denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
            adaptive_thresh = cv2.adaptiveThreshold(denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
            processed_images.append(("adaptive_thresh", adaptive_thresh))
            _, otsu_thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_images.append(("otsu_thresh", otsu_thresh))
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            clahe_img = clahe.apply(gray)
            processed_images.append(("clahe", clahe_img))

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
        
        # Patterns for different fields - made more flexible for OCR variations
        patterns = {
            'PN': [r'P/N\s*[:\s]\s*([A-Z0-9\-\.\/]+)', r'PN\s*[:\s]\s*([A-Z0-9\-\.\/]+)', r'Part\s*No\s*[:\s]\s*([A-Z0-9\-\.\/]+)'],
            'Part_Number': [r'Part\s*Number\s*[:\s]\s*([A-Z0-9\-\.\/]+)', r'Customer\s*Part\s*Number\s*[:\s]\s*([A-Z0-9\-\.\/]+)'],
            'MPN': [r'MPN\s*[:\s]\s*([A-Z0-9\-\.\/]+)', r'MODEL\s*[:\s]\s*([A-Z0-9\-\.\/]+)'],
            'CPN': [r'CPN\s*[:\s]\s*([A-Z0-9\-\.\/]+)'],
            'SSN': [r'SSN\s*[:\s]\s*([A-Z0-9\-\.\/]+)'],
            'SN': [r'S/N\s*[:\s]\s*([A-Z0-9\-\.\/]+)', r'SN\s*[:\s]\s*([A-Z0-9\-\.\/]+)', r'Serial\s*Number\s*[:\s]\s*([A-Z0-9\-\.\/]+)', r'Serial\s*[:\s]\s*([A-Z0-9\-\.\/]+)']
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
        """Look up CPN from MPN in the lookup file"""
        lookup_file = "mpn_cpn_lookup.xlsx"
        if not os.path.exists(lookup_file):
            return None

        try:
            df = pd.read_excel(lookup_file)
            if 'MPN' not in df.columns or 'CPN' not in df.columns:
                print(f"Warning: {lookup_file} must contain 'MPN' and 'CPN' columns.")
                return None

            # Ensure MPN column is string type for matching
            df['MPN'] = df['MPN'].astype(str)

            result = df[df['MPN'] == mpn]

            if not result.empty:
                cpn = result['CPN'].iloc[0]
                return str(cpn)

        except Exception as e:
            print(f"Error reading or looking up in {lookup_file}: {str(e)}")

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
            
    def display_pil_image_in_gui(self, pil_image):
        """Display a PIL image object in the GUI"""
        try:
            # Calculate size to fit in display area
            display_width = 500
            display_height = 300
            image = pil_image.copy()
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
        
    def extract_data_from_current_image(self):
        """Main image processing function, triggered by the 'Extract Data' button."""
        if not self.display_pil_image:
            messagebox.showwarning("No Image", "Please upload an image first.")
            return

        def process():
            try:
                # All heavy processing in background thread
                self.root.after(0, lambda: self.progress.start())
                self.root.after(0, lambda: self.update_status("Processing image..."))
                
                # Preprocess image
                self.root.after(0, lambda: self.update_status("Preprocessing image..."))
                processed_images, original_image = self.preprocess_image(self.display_pil_image)
                
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
                success = self.save_to_excel(self.image_path, extracted_data)
                
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

def main():
    """Main application entry point"""
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()