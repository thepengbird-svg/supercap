#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SUPER CAPT - All-in-One PDF Capture & Crop Tool
User-friendly GUI for fully automated screen capture to PDF creation.
"""

import sys
import os
import time
import threading
import subprocess
from datetime import datetime

# --- Library Installation Check ---
try:
    from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                                   QLabel, QSpinBox, QComboBox, QLineEdit, QFileDialog,
                                   QMessageBox, QProgressBar, QTextEdit, QFrame, QScrollArea)
    from PySide6.QtGui import QFont, QIcon, QScreen, QPainter, QPen, QColor
    from PySide6.QtCore import Qt, Signal, QObject, QRect
    import pyautogui
    from PIL import Image
except ImportError as e:
    missing_lib = str(e).split("'")[1]
    print(f"Required library not found: {missing_lib}")
    print(f"Attempting to install '{missing_lib}'...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", f"{missing_lib}"])
        print("Installation successful. Please restart the application.")
    except Exception as install_e:
        print(f"Failed to install {missing_lib}. Please install it manually: pip install {missing_lib}")
        print(f"Error: {install_e}")
    sys.exit()

# --- Custom Widget for Collapsible Section ---
class CollapsibleBox(QWidget):
    def __init__(self, title="", parent=None):
        super(CollapsibleBox, self).__init__(parent)
        self.toggle_button = QPushButton(f"▼ {title}")
        self.toggle_button.setStyleSheet("QPushButton { text-align: left; border: none; font-weight: bold; }")
        self.toggle_button.setCheckable(True)
        self.toggle_button.setChecked(False)

        self.content_area = QScrollArea()
        self.content_area.setMaximumHeight(0)
        self.content_area.setMinimumHeight(0)
        self.content_area.setWidgetResizable(True)
        self.content_area.setFrameShape(QFrame.NoFrame)

        layout = QVBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.toggle_button)
        layout.addWidget(self.content_area)

        self.toggle_button.toggled.connect(self.toggle)
        self.toggle(False)

    def setContentLayout(self, layout):
        content_widget = QWidget()
        content_widget.setLayout(layout)
        self.content_area.setWidget(content_widget)

    def toggle(self, checked):
        if checked:
            self.toggle_button.setText(f"▲ {self.toggle_button.text().split(' ')[1]}")
            self.content_area.setMaximumHeight(self.content_area.widget().sizeHint().height() + 10)
        else:
            self.toggle_button.setText(f"▼ {self.toggle_button.text().split(' ')[1]}")
            self.content_area.setMaximumHeight(0)

# --- Transparent Overlay for Crop Selection ---
class SnippingWidget(QWidget):
    on_snipped = Signal(QRect)

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen)
        
        self.begin = None
        self.end = None

    def paintEvent(self, event):
        if self.begin and self.end:
            rect = QRect(self.begin, self.end).normalized()
            painter = QPainter(self)
            painter.setPen(QPen(QColor(0, 120, 215, 255), 2))
            painter.setBrush(QColor(0, 120, 215, 50))
            painter.drawRect(rect)

    def mousePressEvent(self, event):
        self.begin = event.pos()
        self.end = event.pos()
        self.update()

    def mouseMoveEvent(self, event):
        self.end = event.pos()
        self.update()

    def mouseReleaseEvent(self, event):
        self.close()
        rect = QRect(self.begin, self.end).normalized()
        self.on_snipped.emit(rect)

# --- Main Application ---
class SuperCaptApp(QWidget):
    class Communicate(QObject):
        log_signal = Signal(str)
        progress_signal = Signal(int, str)
        step_signal = Signal(int, str)
        finished_signal = Signal(str)

    def __init__(self):
        super().__init__()
        self.comm = self.Communicate()
        self.comm.log_signal.connect(self.log)
        self.comm.progress_signal.connect(self.update_progress)
        self.comm.step_signal.connect(self.update_step)
        self.comm.finished_signal.connect(self.process_completed)

        self.crop_area = None
        self.is_capturing = False
        self.is_paused = False
        self.capture_thread = None
        self.snipping_widget = None

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("SUPER CAPT")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setFixedWidth(400)

        # --- Main Layout ---
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # --- Title ---
        title_label = QLabel("SUPER CAPT")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # --- Collapsible Settings ---
        settings_box = CollapsibleBox("Settings")
        main_layout.addWidget(settings_box)
        
        settings_layout = QVBoxLayout()
        
        # Total Pages
        pages_layout = QHBoxLayout()
        pages_layout.addWidget(QLabel("Total Pages:"))
        self.total_pages_spinbox = QSpinBox()
        self.total_pages_spinbox.setRange(1, 9999)
        self.total_pages_spinbox.setValue(200)
        pages_layout.addWidget(self.total_pages_spinbox)
        settings_layout.addLayout(pages_layout)

        # Page Key
        key_layout = QHBoxLayout()
        key_layout.addWidget(QLabel("Page Key:"))
        self.page_key_combo = QComboBox()
        self.page_key_combo.addItems(["Page Down", "Space", "Enter", "Right Arrow", "Down Arrow"])
        self.page_key_combo.setCurrentText("Page Down")
        key_layout.addWidget(self.page_key_combo)
        settings_layout.addLayout(key_layout)

        # Output Folder
        folder_layout = QHBoxLayout()
        self.output_folder_edit = QLineEdit(os.path.join(os.path.expanduser("~"), "Desktop"))
        folder_layout.addWidget(self.output_folder_edit)
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.choose_folder)
        folder_layout.addWidget(self.browse_button)
        settings_layout.addLayout(folder_layout)

        settings_box.setContentLayout(settings_layout)

        # --- Progress Steps ---
        progress_frame = QFrame()
        progress_frame.setFrameShape(QFrame.StyledPanel)
        progress_layout = QVBoxLayout(progress_frame)
        self.step_labels = []
        steps = ["1. Settings", "2. Select Area", "3. Capture", "4. Create PDF", "5. Complete"]
        for step in steps:
            label = QLabel(step)
            label.setStyleSheet("color: gray;")
            self.step_labels.append(label)
            progress_layout.addWidget(label)
        main_layout.addWidget(progress_frame)

        # --- Progress Bar & Label ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("Ready")
        self.progress_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.progress_label)

        # --- Control Buttons ---
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("Start")
        self.start_button.clicked.connect(self.start_process)
        self.start_button.setObjectName("AccentButton")

        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self.pause_process)
        self.pause_button.setEnabled(False)

        self.stop_button = QPushButton("Stop")
        self.stop_button.clicked.connect(self.stop_process)
        self.stop_button.setEnabled(False)

        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.pause_button)
        button_layout.addWidget(self.stop_button)
        main_layout.addLayout(button_layout)
        
        # --- Log Area ---
        log_box = CollapsibleBox("Logs")
        main_layout.addWidget(log_box)
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFixedHeight(100)
        log_layout.addWidget(self.log_text)
        log_box.setContentLayout(log_layout)
        
        self.setLayout(main_layout)
        self.apply_stylesheet()
        self.log("Application started. Check settings and click 'Start'.")

    def apply_stylesheet(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                color: #f0f0f0;
                font-family: Arial;
            }
            QLabel {
                background-color: transparent;
            }
            QPushButton {
                background-color: #4a4a4a;
                border: 1px solid #5a5a5a;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #5a5a5a;
            }
            QPushButton:pressed {
                background-color: #6a6a6a;
            }
            QPushButton#AccentButton {
                background-color: #007acc;
                font-weight: bold;
            }
            QPushButton#AccentButton:hover {
                background-color: #008ae6;
            }
            QSpinBox, QComboBox, QLineEdit {
                background-color: #3c3c3c;
                border: 1px solid #5a5a5a;
                padding: 3px;
                border-radius: 3px;
            }
            QProgressBar {
                border: 1px solid #5a5a5a;
                border-radius: 3px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #007acc;
                border-radius: 3px;
            }
            QTextEdit, QScrollArea {
                background-color: #3c3c3c;
                border: 1px solid #5a5a5a;
                border-radius: 3px;
            }
            QFrame {
                border: 1px solid #4a4a4a;
                border-radius: 3px;
            }
        """)

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")

    def update_step(self, step_index, status):
        colors = {"active": "#007acc", "completed": "#2a9d8f", "pending": "gray"}
        font_weights = {"active": "bold", "completed": "normal", "pending": "normal"}
        
        for i, label in enumerate(self.step_labels):
            if i < step_index:
                label.setStyleSheet(f"color: {colors['completed']}; font-weight: {font_weights['completed']};")
            elif i == step_index:
                label.setStyleSheet(f"color: {colors[status]}; font-weight: {font_weights[status]};")
            else:
                label.setStyleSheet(f"color: {colors['pending']}; font-weight: {font_weights['pending']};")

    def update_progress(self, value, text):
        self.progress_bar.setValue(value)
        self.progress_label.setText(text)

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder", self.output_folder_edit.text())
        if folder:
            self.output_folder_edit.setText(folder)

    def start_process(self):
        # Admin rights check
        try:
            is_admin = os.getuid() == 0
        except AttributeError:
            import ctypes
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        
        if not is_admin:
            QMessageBox.warning(self, "Administrator Rights Required",
                                "For the page-turning feature to work correctly,\n"
                                "please run this application as an administrator.")
            return

        self.log("Starting process...")
        self.comm.step_signal.emit(0, "completed")
        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.stop_button.setEnabled(True)

        QMessageBox.information(self, "Prepare for Capture",
                                "Please ensure:\n\n"
                                "1. The document is open and on the first page.\n"
                                "2. The document window is maximized.\n"
                                "3. Other windows are minimized.\n\n"
                                "Click OK to proceed to select the capture area.")
        
        self.select_crop_area()

    def select_crop_area(self):
        self.log("Selecting crop area...")
        self.comm.step_signal.emit(1, "active")
        self.hide()
        time.sleep(0.5)
        
        self.snipping_widget = SnippingWidget()
        self.snipping_widget.on_snipped.connect(self.on_area_selected)
        self.snipping_widget.show()

    def on_area_selected(self, rect):
        self.show()
        if not rect.isEmpty():
            self.crop_area = (rect.x(), rect.y(), rect.width(), rect.height())
            self.log(f"Crop area selected: {self.crop_area}")
            self.comm.step_signal.emit(1, "completed")
            self.start_capture()
        else:
            self.log("Crop area selection cancelled.")
            self.reset_ui()

    def start_capture(self):
        self.log("Starting automatic capture...")
        self.is_capturing = True
        self.comm.step_signal.emit(2, "active")
        
        self.capture_thread = threading.Thread(target=self.capture_pages_task)
        self.capture_thread.daemon = True
        self.capture_thread.start()

    def get_key(self):
        key_map = {
            "Page Down": "pagedown", "Space": "space", "Enter": "enter",
            "Right Arrow": "right", "Down Arrow": "down"
        }
        return key_map.get(self.page_key_combo.currentText())

    def capture_pages_task(self):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_folder = self.output_folder_edit.text()
            capture_folder = os.path.join(output_folder, f"captured_pages_{timestamp}")
            os.makedirs(capture_folder, exist_ok=True)
            self.comm.log_signal.emit(f"Capture folder created: {capture_folder}")

            self.comm.log_signal.emit("Please click on the document window to focus it. Capture will start in 3 seconds...")
            time.sleep(3)

            total = self.total_pages_spinbox.value()
            page_key = self.get_key()

            for page in range(1, total + 1):
                if not self.is_capturing:
                    self.comm.log_signal.emit("Capture stopped by user.")
                    return

                while self.is_paused:
                    time.sleep(0.1)

                self.comm.log_signal.emit(f"Capturing page {page}/{total}...")
                
                screenshot = pyautogui.screenshot()
                filepath = os.path.join(capture_folder, f"page_{page:04d}.png")
                screenshot.save(filepath)

                progress = int((page / total) * 50) # Capture is first 50%
                self.comm.progress_signal.emit(progress, f"Captured: {page}/{total}")
                
                if page < total:
                    pyautogui.press(page_key)
                    time.sleep(1.5) # Wait for page to load

            if self.is_capturing:
                self.comm.log_signal.emit("Capture phase complete.")
                self.comm.step_signal.emit(2, "completed")
                self.crop_and_create_pdf_task(capture_folder, timestamp)

        except Exception as e:
            self.comm.log_signal.emit(f"Error during capture: {e}")
            self.reset_ui()

    def crop_and_create_pdf_task(self, capture_folder, timestamp):
        try:
            self.comm.log_signal.emit("Cropping and creating PDF...")
            self.comm.step_signal.emit(3, "active")
            
            output_folder = self.output_folder_edit.text()
            cropped_folder = os.path.join(output_folder, f"cropped_pages_{timestamp}")
            os.makedirs(cropped_folder, exist_ok=True)

            image_files = sorted([f for f in os.listdir(capture_folder) if f.endswith('.png')])
            total_images = len(image_files)

            for i, filename in enumerate(image_files, 1):
                if not self.is_capturing:
                    self.comm.log_signal.emit("Process stopped during cropping.")
                    return

                input_path = os.path.join(capture_folder, filename)
                output_path = os.path.join(cropped_folder, filename)
                
                img = Image.open(input_path)
                x, y, w, h = self.crop_area
                cropped = img.crop((x, y, x + w, y + h))
                cropped.save(output_path)
                
                progress = 50 + int((i / total_images) * 40) # Cropping is 50-90%
                self.comm.progress_signal.emit(progress, f"Cropping: {i}/{total_images}")

            self.comm.log_signal.emit("Cropping complete. Creating PDF...")
            self.comm.step_signal.emit(3, "completed")
            self.comm.step_signal.emit(4, "active")
            self.comm.progress_signal.emit(95, "Creating PDF...")

            # Create PDF
            pdf_path = os.path.join(output_folder, f"SUPER_CAPT_{timestamp}.pdf")
            cropped_files = sorted([os.path.join(cropped_folder, f) for f in os.listdir(cropped_folder) if f.endswith('.png')])
            
            if not cropped_files:
                raise ValueError("No cropped images found to create PDF.")

            first_image = Image.open(cropped_files[0]).convert("RGB")
            other_images = [Image.open(f).convert("RGB") for f in cropped_files[1:]]
            
            first_image.save(pdf_path, save_all=True, append_images=other_images, resolution=150.0)
            
            self.comm.log_signal.emit(f"PDF created successfully: {pdf_path}")
            self.comm.finished_signal.emit(pdf_path)

        except Exception as e:
            self.comm.log_signal.emit(f"Error during PDF creation: {e}")
            self.reset_ui()

    def pause_process(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.setText("Resume")
            self.log("Process paused.")
        else:
            self.pause_button.setText("Pause")
            self.log("Process resumed.")

    def stop_process(self):
        if QMessageBox.question(self, "Confirm Stop", "Are you sure you want to stop the process?") == QMessageBox.Yes:
            self.is_capturing = False
            self.log("Process stopping...")
            self.reset_ui()

    def reset_ui(self):
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.pause_button.setText("Pause")
        self.stop_button.setEnabled(False)
        self.update_progress(0, "Ready")
        self.update_step(-1, "pending")
        self.is_capturing = False
        self.is_paused = False

    def process_completed(self, pdf_path):
        self.reset_ui()
        self.comm.progress_signal.emit(100, "Complete!")
        self.comm.step_signal.emit(4, "completed")

        reply = QMessageBox.information(self, "Success!",
                                        f"PDF created successfully!\n\nPath: {pdf_path}",
                                        QMessageBox.Ok | QMessageBox.Open)
        
        if reply == QMessageBox.Open:
            try:
                os.startfile(pdf_path)
            except Exception as e:
                self.log(f"Failed to open PDF: {e}")

    def closeEvent(self, event):
        if self.is_capturing:
            reply = QMessageBox.question(self, 'Confirm Exit',
                                         "A capture process is running. Are you sure you want to exit?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.is_capturing = False
                if self.capture_thread:
                    self.capture_thread.join(timeout=2) # Wait briefly for thread to exit
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SuperCaptApp()
    ex.show()
    sys.exit(app.exec())
