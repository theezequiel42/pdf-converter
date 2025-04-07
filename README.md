# PDF Converter: A Versatile GUI Tool for PDF and Image Manipulation

This project provides a user-friendly graphical interface for performing various operations on PDF files and images. It offers three main functionalities: dividing PDFs, combining images into PDFs, and extracting images from PDFs.

The PDF Converter is designed to simplify common document management tasks, making it easier for users to work with PDF files and images without the need for complex software. Whether you need to extract specific pages from a PDF, create a PDF from a collection of images, or convert a PDF into individual image files, this tool has you covered.

## Repository Structure

The repository contains a single Python script:

- `pdf_converter.py`: The main script that implements the GUI and all PDF/image conversion functionalities.

## Usage Instructions

### Installation

1. Ensure you have Python 3.6 or later installed on your system.
2. Install the required dependencies:

```bash
pip install PyPDF2 Pillow PyMuPDF PyQt6 pdf2docx
```

### Getting Started

To run the PDF Converter, execute the following command in your terminal:

```bash
python pdf_converter.py
```

This will launch the graphical user interface with buttons for each available function.

### Features and How to Use Them

1. **Divide PDF**
   - Click the "Divide PDF" button.
   - Select the PDF file you want to divide.
   - Enter the page numbers you want to extract (e.g., "1,2,5-7").
   - The selected pages will be saved as a new PDF file.

2. **Images to PDF**
   - Click the "Images to PDF" button.
   - Select multiple image files (PNG, JPG, JPEG).
   - Choose a location to save the resulting PDF.
   - The images will be combined into a single PDF file.

3. **PDF to Images**
   - Click the "PDF to Images" button.
   - Select the PDF file you want to convert.
   - Choose a directory to save the extracted images.
   - Each page of the PDF will be saved as a separate PNG image.

4. **PDF to Word**
   - Click the "PDF to Word" button.
   - Select the PDF file you want to convert.
   - Choose a location to save the resulting DOCX file.

5. **Merge PDFs**
   - Click the "Merge PDFs" button.
   - Select multiple PDF files you want to combine.
   - Choose a location to save the merged PDF.

6. **Protect PDF**
   - Click the "Protect PDF with Password" button.
   - Select the PDF file.
   - Enter a password to protect the document.
   - Choose a location to save the protected PDF.

7. **Compress PDF**
   - Click the "Compress PDF" button.
   - Select the PDF file you want to compress.
   - Choose a location to save the compressed version.

### Troubleshooting

- If you encounter a "No module named" error, ensure you've installed all required dependencies.
- For "Permission denied" errors, check that you have write access to the output directory.
- If images are not displaying correctly, verify that the input files are valid image formats.

## Data Flow

The PDF Converter processes data in the following manner:

1. User input is received through the GUI.
2. Input files (PDFs or images) are read using PyPDF2 or PIL.
3. Data is processed according to the selected function.
4. Output files are generated and saved to the user-specified location.
5. Feedback is provided to the user via message boxes.

```
[User Input] -> [GUI] -> [File Selection] -> [Data Processing] -> [File Output] -> [User Feedback]
```

Note: The application handles files locally and does not involve any network operations or external services.