# PDF to Excel Converter

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Description

Convert PDF documents into Excel files effortlessly with this powerful PDF to Excel converter built using Python and Flask.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Demo](#demo)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgements](#acknowledgements)

## Introduction

The PDF to Excel Converter is a web application developed using Flask, a micro web framework for Python. This application enables users to upload PDF files and convert them into Excel spreadsheets. It utilizes the pdfplumber and tabula-py libraries to extract textual content and tables from PDFs, then formats and exports them into Excel files.

## Features

- Convert entire PDF text content into Excel format.
- Extract tables from PDF pages and export them as separate sheets in Excel.
- Easy-to-use web interface for uploading PDF files.
- Supports multiple processing options for different user needs.
- Well-structured codebase for easy customization and extension.

## Demo

[Insert demo video or link to live demo if available]

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/yourusername/pdf-to-excel-converter.git
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
    python app.py
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

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [pdfplumber](https://github.com/jsvine/pdfplumber)
- [tabula-py](https://github.com/chezou/tabula-py)
- [Flask](https://flask.palletsprojects.com/)

