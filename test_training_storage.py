import os
import sqlite3
import shutil
from pathlib import Path
from unittest.mock import patch
from main import OCRApp
import tkinter as tk
import datetime

class MockBooleanVar:
    """A mock class to simulate tk.BooleanVar without a root window."""
    def __init__(self, value=False):
        self._value = value
    def get(self):
        return self._value
    def set(self, value):
        self._value = value

class HeadlessOCRApp(OCRApp):
    """A version of OCRApp that doesn't initialize the GUI, for testing."""
    def __init__(self):
        # We need a mock root for messagebox calls, but we can't create one.
        # We will mock the messageboxes themselves in the test.
        self.root = None

        self.deep_scan_enabled = MockBooleanVar(value=False)
        self.db_file = "test_final_database.db"
        self.training_data_dir = "test_final_training_data"

        if os.path.exists(self.db_file):
            os.remove(self.db_file)
        if os.path.exists(self.training_data_dir):
            shutil.rmtree(self.training_data_dir)

        Path(self.training_data_dir).mkdir(exist_ok=True)
        super().init_database()

    def finish_training_session(self):
        """Save all collected labels from the session to the database."""
        if not hasattr(self, 'training_files') or not self.training_files:
            return

        saved_count, skipped_count = 0, 0
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            for i, file_path in enumerate(self.training_files):
                labels = self.training_labels[i]
                if not any(labels.values()):
                    continue
                filename = os.path.basename(file_path)
                dest_path = os.path.join(self.training_data_dir, filename)
                shutil.copy2(file_path, dest_path)
                db_path = dest_path.replace('\\', '/')
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                try:
                    cursor.execute('''
                        INSERT OR REPLACE INTO training_data
                        (image_path, pn, part_number, mpn, cpn, ssn, sn, timestamp)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (db_path, labels.get('PN',''), labels.get('Part_Number',''), labels.get('MPN',''), labels.get('CPN',''), labels.get('SSN',''), labels.get('SN',''), timestamp))
                    saved_count += 1
                except sqlite3.Error as e:
                    print(f"DB Error for {filename}: {e}")
                    skipped_count += 1
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Failed to save training session: {e}")

def run_final_test():
    print("--- Setting up final test environment ---")
    app = HeadlessOCRApp()

    dummy_image_name = "dummy_image_for_training.jpg"
    dummy_source_path = "dummy_image_for_training.jpg"
    with open(dummy_source_path, "w") as f:
        f.write("dummy data")

    print("\n--- Testing Training Data Storage ---")
    app.training_files = [dummy_source_path]
    app.current_training_index = 0
    app.training_labels = [{'MPN': 'FINAL-MPN', 'CPN': 'FINAL-CPN', 'PN': 'FINAL-PN', 'Part_Number': '', 'SSN': '', 'SN': 'FINAL-SN'}]

    with patch('tkinter.messagebox.showinfo'), \
         patch('tkinter.messagebox.showwarning'), \
         patch('tkinter.messagebox.showerror'):
        app.finish_training_session()

    # Verify the database content
    conn = sqlite3.connect(app.db_file)
    cursor = conn.cursor()
    cursor.execute("SELECT mpn, cpn, sn FROM training_data WHERE image_path LIKE ?", (f'%{dummy_image_name}',))
    result = cursor.fetchone()
    conn.close()

    print(f"  Querying database for saved training data... Found: {result}")
    assert result is not None, "Training data was not saved to the database."
    assert result[0] == 'FINAL-MPN', f"MPN mismatch. Expected FINAL-MPN, got {result[0]}"
    assert result[1] == 'FINAL-CPN', f"CPN mismatch. Expected FINAL-CPN, got {result[1]}"
    assert result[2] == 'FINAL-SN', f"SN mismatch. Expected FINAL-SN, got {result[2]}"
    print("  Final training data storage test PASSED.")

    print("\n--- Cleaning up test files ---")
    if os.path.exists(app.db_file):
        os.remove(app.db_file)
    if os.path.exists(app.training_data_dir):
        shutil.rmtree(app.training_data_dir)
    if os.path.exists(dummy_source_path):
        os.remove(dummy_source_path)
    print("All tests passed and cleanup complete.")

if __name__ == "__main__":
    run_final_test()
