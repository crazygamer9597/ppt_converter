# 📄 Office File to PDF Converter

A robust, user-friendly Python tool for batch converting Microsoft Office documents (Word and PowerPoint) to PDF format with beautiful progress visualization and comprehensive logging.

## 🚀 Features

- **Batch Conversion**: Convert multiple Office documents to PDF in one go
- **Smart Processing**: Automatically skips files that have already been converted
- **Beautiful UI**: Rich progress bars and colorful console output using the `rich` library
- **Comprehensive Logging**: Detailed logs saved to file for troubleshooting
- **Multiple Formats**: Supports Word documents (.doc, .docx) and PowerPoint presentations (.ppt, .pptx)
- **PDF Copying**: Automatically copies existing PDF files to the output directory
- **Flexible Usage**: Both interactive and command-line modes
- **Error Handling**: Robust error handling with detailed error messages
- **Performance Optimized**: Efficient COM application management and resource cleanup

## 🛠️ Installation

### 1. Clone the Repository

```bash
https://github.com/crazygamer9597/ppt_converter.git
cd ppt_converter
```

### 2. Install Required Packages

```bash
pip install -r requirements.txt
```

**Or install packages individually:**

```bash
pip install comtypes rich
```

### 3. Verify Installation

```bash
python file_converter.py --version
```

### Package Details

- **`comtypes`**: Provides COM automation support for interacting with Microsoft Office applications
- **`rich`**: Creates beautiful terminal output with progress bars, tables, and colored text

## 🎯 Usage

### Interactive Mode

Simply run the script without arguments for an interactive experience:

```bash
python file_converter.py
```

The program will prompt you to enter the input directory path.

### Command Line Mode

#### Basic Usage

```bash
python file_converter.py /path/to/your/documents
```

#### Specify Output Directory

```bash
python file_converter.py /path/to/input -o /path/to/output
```

#### Custom Log File

```bash
python file_converter.py /path/to/input --log-file my_conversion.log
```

#### Get Help

```bash
python file_converter.py --help
```

## 📁 File Organization

### Input Directory Structure

```
input_folder/
├── document1.docx
├── presentation1.pptx
├── old_document.doc
├── slides.ppt
├── existing.pdf
└── other_files.txt
```

### Output Directory Structure

```
input_folder/converted_pdf/
├── document1.pdf
├── presentation1.pdf
├── old_document.pdf
├── slides.pdf
└── existing.pdf (copied)
```

## ⚙️ Configuration

The tool uses a configuration class that can be customized:

- **Supported formats**: .doc, .docx, .ppt, .pptx
- **Output format**: PDF
- **Default output folder**: `converted_pdf` (created in input directory)
- **Default log file**: `conversion_log.txt`

## 📊 Features in Detail

### Progress Visualization

- Real-time progress bar showing current file being processed
- Percentage completion and estimated time remaining
- File count progress (e.g., "3/10 files")

### Smart Processing

- Skips files that have already been converted to avoid duplicates
- Copies existing PDF files to maintain complete document sets
- Handles various Office document formats seamlessly

### Comprehensive Logging

- Detailed timestamps for all operations
- Success and error messages
- File-specific processing information
- Separate log file for easy troubleshooting

### Error Handling

- Graceful handling of corrupted files
- COM application error recovery
- User-friendly error messages
- Continues processing even if individual files fail

## 🔧 Troubleshooting

### Common Issues

#### "No module named 'comtypes'"

```bash
pip install comtypes
```

#### "No module named 'rich'"

```bash
pip install rich
```

#### COM Application Errors

- Ensure Microsoft Office is properly installed
- Try running the script as administrator
- Close all Office applications before running the converter

#### Permission Errors

- Ensure the output directory is writable
- Run the script as administrator if necessary
- Check that input files are not currently open in Office applications

### Debug Mode

Check the log file (`conversion_log.txt` by default) for detailed error information.

## 🙏 Acknowledgments

- [Rich](https://github.com/Textualize/rich) - For beautiful terminal output
- [comtypes](https://github.com/enthought/comtypes) - For COM automation support
- Microsoft Office - For providing the COM interfaces

## 📈 Version History

- **v2.0.0**: Complete rewrite with improved architecture, command-line support, and better error handling
- **v1.0.0**: Initial release with basic conversion functionality

## 🐛 Bug Reports

Please report bugs by creating an issue on GitHub with:

- Operating system version
- Python version
- Microsoft Office version
- Error message or log file content
- Steps to reproduce the issue

## 💡 Feature Requests

We welcome feature requests! Please create an issue describing:

- The feature you'd like to see
- Why it would be useful

---

**Made with ❤️ for efficient document conversion**
