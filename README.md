# PDF Editor

A Python-based PDF editor application with a Tkinter GUI. This application provides functionalities to merge, split, rotate, delete pages from PDFs, and convert PDFs to Word documents.

## Features

- **Merge PDFs**: Combine multiple PDF files into one.
- **Delete Page**: Remove specific pages from a PDF file.
- **Split PDF**: Split a PDF file into two separate files at a specified page.
- **Rotate Page**: Rotate a specific page in a PDF by 90, 180, or 270 degrees.
- **Convert PDF to Word**: Convert a PDF file to a Word document (.docx).

## Requirements

- Python 3.6+
- Required Python libraries (see `requirements.txt`)

## Installation

1. **Clone the repository**:

    ```bash
    git clone https://github.com/your-username/pdf_editor.git
    cd pdf_editor
    ```

2. **Set up a virtual environment**:
    - On Windows:

      ```bash
      python -m venv env
      .\env\Scripts\activate
      ```

    - On macOS/Linux:

      ```bash
      python3 -m venv env
      source env/bin/activate
      ```

3. **Install the required libraries**:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. **Run the application**:

    ```bash
    python main.py
    ```

2. **Using the GUI**:
    - **Merge PDFs**: Click "Merge PDFs", select the PDF files to merge, choose the save location, and save the merged PDF.
    - **Delete Page from PDF**: Click "Delete Page from PDF", select the PDF file, enter the page number to delete, choose the save location, and save the edited PDF.
    - **Split PDF**: Click "Split PDF", select the PDF file, enter the page number to split at, choose the save locations for both parts, and save the split PDFs.
    - **Rotate Page in PDF**: Click "Rotate Page in PDF", select the PDF file, enter the page number to rotate, specify the rotation angle, choose the save location, and save the rotated PDF.
    - **Convert PDF to Word**: Click "Convert PDF to Word", select the PDF file, choose the save location, and save the Word document.

## Building an Executable

If you want to create a standalone executable that can run on another PC without requiring Python to be installed, follow these steps:

1. **Install PyInstaller**:

    ```bash
    pip install pyinstaller
    ```

2. **Generate the executable**:

    ```bash
    pyinstaller --onefile --windowed main.py
    ```

3. **Find the executable**:
    - The executable will be located in the `dist` directory.

## Contributing

1. Fork the repository.
2. Create a new branch: `git checkout -b my-feature-branch`.
3. Make your changes and commit them: `git commit -m 'Add some feature'`.
4. Push to the branch: `git push origin my-feature-branch`.
5. Submit a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the developers of the libraries used in this project.
- Special thanks to all contributors.
