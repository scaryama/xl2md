"""
GUI application for converting XLSX files to Markdown format
Uses PyQt6 to accept files via drag and drop and convert them.

Required packages:
    pip install PyQt6 pandas openpyxl
"""

import sys
import os
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QPushButton, QFileDialog, QMessageBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont

# Import existing conversion module
from xl2md import xlsx_to_markdown


class ConversionThread(QThread):
    """Thread that runs conversion tasks in the background"""
    log_signal = pyqtSignal(str)  # Signal to pass log messages
    finished_signal = pyqtSignal(bool)  # Signal to pass completion status
    
    def __init__(self, xlsx_path, output_dir=None):
        super().__init__()
        self.xlsx_path = xlsx_path
        self.output_dir = output_dir
    
    def run(self):
        """Conversion task to be executed in the thread"""
        def log_callback(message):
            """Callback function: passes messages via signal"""
            self.log_signal.emit(message)
        
        try:
            success = xlsx_to_markdown(
                self.xlsx_path,
                self.output_dir,
                log_callback=log_callback
            )
            self.finished_signal.emit(success)
        except Exception as e:
            self.log_signal.emit(f"❌ Thread error: {e}")
            import traceback
            self.log_signal.emit(traceback.format_exc())
            self.finished_signal.emit(False)


class DropArea(QWidget):
    """Drag and drop area widget"""
    file_dropped = pyqtSignal(str)  # Signal for file drop
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(100)
        self.setStyleSheet("""
            DropArea {
                border: 3px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                padding: 20px;
            }
            DropArea:hover {
                border-color: #0078d4;
                background-color: #e8f4f8;
            }
        """)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.label = QLabel("Drag and drop XLSX file here")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("font-size: 14px; color: #666;")
        layout.addWidget(self.label)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Drag enter event"""
        if event.mimeData().hasUrls():
            # Only allow XLSX files
            urls = event.mimeData().urls()
            if urls and any(url.toLocalFile().lower().endswith('.xlsx') for url in urls):
                event.acceptProposedAction()
                self.setStyleSheet("""
                    DropArea {
                        border: 3px dashed #0078d4;
                        border-radius: 10px;
                        background-color: #e8f4f8;
                        padding: 20px;
                    }
                """)
            else:
                event.ignore()
        else:
            event.ignore()
    
    def dragLeaveEvent(self, event):
        """Drag leave event"""
        self.setStyleSheet("""
            DropArea {
                border: 3px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                padding: 20px;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        """Drop event"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            xlsx_files = [url.toLocalFile() for url in urls if url.toLocalFile().lower().endswith('.xlsx')]
            
            if xlsx_files:
                # Use only the first XLSX file
                file_path = xlsx_files[0]
                self.file_dropped.emit(file_path)  # Emit signal
                event.acceptProposedAction()
            else:
                event.ignore()
        
        # Restore style
        self.setStyleSheet("""
            DropArea {
                border: 3px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                padding: 20px;
            }
        """)


class MainWindow(QMainWindow):
    """Main window"""
    
    def __init__(self):
        super().__init__()
        self.conversion_thread = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize UI"""
        self.setWindowTitle("XLSX to Markdown Converter")
        self.setGeometry(100, 100, 800, 600)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Title
        title_label = QLabel("XLSX to Markdown Converter")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)
        
        # Drag and drop area
        self.drop_area = DropArea()
        self.drop_area.file_dropped.connect(self.handle_file_drop)  # Connect signal
        main_layout.addWidget(self.drop_area)
        
        # File selection button
        button_layout = QHBoxLayout()
        self.file_button = QPushButton("Select File")
        self.file_button.clicked.connect(self.select_file)
        button_layout.addWidget(self.file_button)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)
        
        # Log area
        log_label = QLabel("Progress:")
        log_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        main_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        main_layout.addWidget(self.log_text)
        
        # Status label
        self.status_label = QLabel("Waiting...")
        self.status_label.setStyleSheet("padding: 5px; background-color: #f0f0f0;")
        main_layout.addWidget(self.status_label)
        
        # Initial message
        self.log_message("=" * 60)
        self.log_message("XLSX to Markdown Converter (Separated by Sheet)")
        self.log_message("=" * 60)
        self.log_message("Drag and drop an XLSX file or click the 'Select File' button.\n")
    
    def log_message(self, message):
        """Add log message to text area"""
        self.log_text.append(message)
        # Auto scroll
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def select_file(self):
        """File selection dialog"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select XLSX File",
            "",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            self.start_conversion(file_path)
    
    def handle_file_drop(self, file_path):
        """Handle dropped file"""
        self.start_conversion(file_path)
    
    def start_conversion(self, xlsx_path):
        """Start conversion"""
        # Abort if conversion is already in progress
        if self.conversion_thread and self.conversion_thread.isRunning():
            QMessageBox.warning(self, "Notice", "Conversion is already in progress.")
            return
        
        # Check if file exists
        if not os.path.exists(xlsx_path):
            QMessageBox.critical(self, "Error", f"File not found:\n{xlsx_path}")
            return
        
        # Initialize log
        self.log_text.clear()
        self.log_message("=" * 60)
        self.log_message("XLSX to Markdown Converter (Separated by Sheet)")
        self.log_message("=" * 60)
        self.log_message(f"\nFile: {xlsx_path}\n")
        
        # Update status
        self.status_label.setText("Converting...")
        self.status_label.setStyleSheet("padding: 5px; background-color: #fff3cd;")
        
        # Disable buttons
        self.file_button.setEnabled(False)
        self.drop_area.setEnabled(False)
        
        # Start conversion thread
        self.conversion_thread = ConversionThread(xlsx_path)
        self.conversion_thread.log_signal.connect(self.log_message)
        self.conversion_thread.finished_signal.connect(self.on_conversion_finished)
        self.conversion_thread.start()
    
    def on_conversion_finished(self, success):
        """Handle conversion completion"""
        # Enable buttons
        self.file_button.setEnabled(True)
        self.drop_area.setEnabled(True)
        
        if success:
            self.status_label.setText("Conversion Complete!")
            self.status_label.setStyleSheet("padding: 5px; background-color: #d4edda;")
            self.log_message("\n✨ Conversion completed successfully!")
        else:
            self.status_label.setText("Conversion Failed")
            self.status_label.setStyleSheet("padding: 5px; background-color: #f8d7da;")
            self.log_message("\n❌ Conversion failed.")
        
        # Clean up thread
        self.conversion_thread = None


def main():
    """Main function"""
    app = QApplication(sys.argv)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

