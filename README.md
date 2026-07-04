[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

# Latest Update v0.2.0-demo -> v0.2.0
## Release Notes 25/06/2026:
- Added Auto-shutdown when browser closes (heartbeat mechanism to auto-shutdown server when browser closes)
- Added Multi-file upload with ZIP download
- Added CSV and JSON export formats
- Added Drag & drop file upload
- Added Desktop [installer for Windows](https://github.com/TsvetanG2/PDF-To-Excel-Converter/releases/download/v0.2.0/PDF-To-Excel-Converter-Setup-0.2.0.exe)
- Nested server loops fix
- Bug fixes
- Increased speed and added multi-threading + logging
- Added Rate limiting (10 req/min per IP)
- Added Dockerfile for containerization


# PDF to Excel Converter

## Introduction

<img width="754" height="631" alt="download" src="https://github.com/user-attachments/assets/146f145a-5bd7-45b8-9831-c300a0210c52" />

#

The pdf-to-excel-converter repository provides a web application designed to convert PDF documents into Excel spreadsheets. This application allows users to upload PDF files through a web interface and select specific conversion options, ultimately enabling them to download the processed data as an Excel file.

The core functionality of the application is centered around two primary extraction modes:

### All Text + Tables: 
This mode extracts all textual content and tabular data from the PDF and consolidates it into a single Excel worksheet. This process involves merging diverse data types into a coherent output.
### Tables Only: 
This mode focuses exclusively on identifying and extracting tabular data from the PDF. Each detected table is then organized into a separate sheet within the generated Excel workbook. This is particularly useful for PDFs where structured data in tables is the main interest.
Users interact with the application by navigating to the main page, where they can upload a PDF file and choose their preferred extraction mode. Upon submission, the application processes the PDF, extracts the relevant content, generates an Excel file, and then offers it for download. The web application handles the file upload, initiates the conversion process, and manages the delivery of the final Excel output. The overall architecture is built around a Flask web application, which orchestrates these operations. For more details on the web application's structure and workflow, see Flask Application Structure and Workflow.

# Design / UI / Screenshot

<img width="690" height="700" alt="Screenshot 2026-07-04 030341" src="https://github.com/user-attachments/assets/783f7315-92e6-4288-ba89-385175bec40f" />


## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgements](#acknowledgements)

## Features

- Convert entire PDF text content into Excel format.
- Extract tables from PDF pages and export them as separate sheets in Excel.
- Easy-to-use web interface for uploading PDF files.
- Supports multiple processing options for different user needs.
- Well-structured codebase for easy customization and extension.

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/TsvetanG2/pdf-to-excel-converter.git
    ```

2. Navigate to the project directory:

    ```bash
    cd pdf-to-excel-converter
    ```

3. Install dependencies:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Start the Flask server:

    ```bash
    python pdftoexcel.py
    ```

2. Open your web browser and navigate to http://localhost:5000.

3. Upload a PDF file using the provided form.

4. Choose the processing option:
    - All Text: Convert entire text content of the PDF into an Excel file.
    - Tables Only: Extract tables from the PDF and export them as separate sheets in Excel.

5. Click on the "Convert" button and wait for the conversion to complete.

6. Download the generated Excel file.

## Contributing

Contributions are welcome! Please follow these steps to contribute:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/yourfeature`).
3. Make your changes.
4. Commit your changes (`git commit -am 'Add new feature'`).
5. Push to the branch (`git push origin feature/yourfeature`).
6. Create a new Pull Request.

- **Documentation**: [docs/](docs/) folder
- **Bug Reports**: [GitHub Issues](https://github.com/TsvetanG2/PDF-To-Excel-Converter/issues)
- **Feature Requests**: [GitHub Discussions](https://github.com/TsvetanG2/PDF-To-Excel-Converter/discussions)
- **Show Support**: Star the repository if you find it useful!

## License

This project is licensed under the MIT License - see the [LICENSE](https://opensource.org/license/MIT) file for details.

## Acknowledgements

- [pdfplumber](https://github.com/jsvine/pdfplumber)
- [tabula-py](https://github.com/chezou/tabula-py)
- [Flask](https://flask.palletsprojects.com/)

