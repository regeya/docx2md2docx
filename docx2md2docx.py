#!/usr/bin/env python

# This script uses Pyside6 for file dialogs, pandoc for the heavy lifting, and Neovim for the editing.
# Run this script in Windows WSL.

import subprocess

from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog
from pathlib import Path
import sys
import os
import tempfile


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.default_directory = "/mnt/c/Users/shane/Downloads"
        self.relative_filename = ""
        self.new_filename = ""
        self.old_filename = ""

        # Open file dialog automatically when the main window initializes
        self.open_file_dialog()
        self.tmp_file_fd, self.temp_file_path = tempfile.mkstemp()
        os.close(self.tmp_file_fd)
        self.temp_file_path = self.temp_file_path + ".md"
        self.convert_file(self.old_filename, self.temp_file_path)
        self.process = subprocess.Popen(
            ["/bin/kitty nvim {}".format(self.temp_file_path)], shell=True
        )
        self.process.wait()
        self.open_save_dialog()
        self.convert_file(self.temp_file_path, self.new_filename)

    def default_new_filename(self):
        name, extension = os.path.splitext(self.relative_filename)
        newname = "{}-converted".format(name)
        full_filename = Path(f"{newname}{extension}")
        full_path = self.default_directory / full_filename
        return str(full_path)

    def open_file_dialog(self):
        # Open a file dialog and get the file path
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open File",
            self.default_directory,
            "Word 2010 Files (*.docx);;All Files (*)",
        )
        if file_path:
            self.relative_filename = os.path.basename(file_path)
            self.old_filename = file_path
        else:
            print("No file selected.")

    def open_save_dialog(self):
        # Open a file dialog and get the file path
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Open File",
            self.default_new_filename(),
            "Word 2010 Files (*.docx);;All Files (*)",
        )
        if file_path:
            self.new_filename = str(file_path)
        else:
            print("No file selected.")

    def convert_file(self, input_file, output_file):
        try:
            # Run Pandoc command
            subprocess.run(
                ["pandoc", input_file, "-o", output_file],
                check=True,  # Raises an error if the command fails
            )
            print(f"Converted {input_file} to {output_file}")
        except subprocess.CalledProcessError as e:
            print("Error occurred:", e)


app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec())
