# DOCX Image Converter Tool - README

## Overview
This tool converts images in `.docx` files from formats like PNG, BMP, TIFF, and WebP to optimized JPGs. It also resizes the images based on a specified DPI, helping reduce the overall file size of the document while maintaining good image quality.

## Usage

### 1. Input File Path
- You can input either the **full file path** (e.g., `C:/Documents/myfile.docx`) or a **relative path** (e.g., `uploads/myfile.docx`).

### 2. Configuration
The image quality and DPI settings are controlled through the `settings.cfg` file.

- **Quality**: Ranges from **1** (low) to **100** (high).
- **DPI**: Controls the image resolution. Common values are around **150-200 DPI** for good quality.

### 3. Running the Program
Run the script and provide the path to your `.docx` file when prompted.

```bash
python convert_docx_images.py
```

The converted file will be saved in the `output` folder with the prefix `converted_`.

### 4. File Size Reduction
After conversion, the program will display the file size reduction percentage.

## Conclusion
This tool simplifies `.docx` image conversion, improving document efficiency while keeping image quality intact. Adjust the settings to suit your quality and file size preferences.