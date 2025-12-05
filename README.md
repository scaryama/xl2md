# 📊 XLSX to Markdown Converter

A tool that automatically converts each sheet of Excel files (.xlsx) into individual Markdown files. Available in both GUI and command-line versions.

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![PyQt6](https://img.shields.io/badge/PyQt6-6.0+-green.svg)](https://www.riverbankcomputing.com/static/Docs/PyQt6/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## ✨ Features

- 📄 **Excel File Reading**: Automatically recognizes all sheets in .xlsx format files
- 📝 **Markdown Conversion**: Converts each sheet into individual Markdown tables
- 🖱️ **Drag and Drop**: Easily convert files by dragging them in the GUI
- 📊 **Real-time Progress**: Monitor conversion progress in real-time
- 🔄 **Asynchronous Processing**: GUI remains responsive during conversion
- 📦 **Executable Support**: Generate .exe files that run without Python installation

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd python

# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt
```

### Usage

#### GUI Version (Recommended)

```bash
python xl2md_gui.py
```

1. When the GUI window opens, drag and drop an XLSX file or click the "Select File" button
2. Monitor the conversion progress
3. After conversion completes, Markdown files are created in the same location as the source file

#### Command-Line Version

```bash
# Basic usage
python xl2md.py

# Specify a file
python xl2md.py "path/to/file.xlsx"

# Specify output directory
python xl2md.py "path/to/file.xlsx" "output/path"
```

## 📋 Requirements

- Python 3.8 or higher
- Windows 10 or higher (GUI version)

### Required Packages

```
pandas>=1.3.0
openpyxl>=3.0.0
PyQt6>=6.0.0
```

## 📁 Project Structure

```
python/
├── xl2md.py              # Command-line version (core conversion logic)
├── xl2md_gui.py          # GUI version
├── build_exe.bat         # Windows batch build script
├── build_exe.ps1         # PowerShell build script
├── README_BUILD.md       # Build guide
└── doc/                  # Excel files to convert
```

## 🔧 Build and Deployment

### Creating Executable (.exe) File

```bash
# Using batch file
build_exe.bat

# Or using PowerShell script
.\build_exe.ps1
```

After build completes, the `dist\XLSX2Markdown.exe` file will be created.

For detailed build instructions, please refer to [README_BUILD.md](README_BUILD.md).

## 📖 Usage Examples

### Input File
- `example.xlsx` (contains multiple sheets)

### Output Files
Individual Markdown files are created for each sheet:
- `example_Sheet1.md`
- `example_Sheet2.md`
- ...

### Markdown File Format

```markdown
# Sheet Name
*Source File: example.xlsx*
*Sheet Name: Sheet1*
---

| Column1 | Column2 | Column3 |
|---------|---------|---------|
| Data1    | Data2   | Data3   |
```

## 🛠️ Tech Stack

- **Python 3.8+**: Programming language
- **PyQt6**: GUI framework
- **pandas**: Excel file reading and data processing
- **openpyxl**: Excel file parsing engine
- **PyInstaller**: Executable build tool

## 📝 Key Features Details

### Filename Processing
- Automatically removes characters that cannot be used in Windows
- Converts spaces to underscores
- Prevents filename conflicts

### Special Character Handling
- Pipe character (`|`): Escaped as `\|`
- Line breaks: Converted to `<br>` tags
- NaN values: Converted to empty strings

### Error Handling
- File existence verification
- Empty sheet processing
- Continues processing even if individual sheet processing fails
- Provides detailed error messages and tracebacks

## 🤝 Contributing

Bug reports, feature suggestions, and Pull Requests are welcome!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is distributed under the MIT License.

## 🙏 Acknowledgments

This project was developed to convert Excel-based planning documents into Markdown format, making it easier to integrate with version control systems.

## 📚 Related Documentation

- [Build Guide](README_BUILD.md)
- [Development Documentation](README.md)

---

**Made with ❤️ for better documentation workflow**

