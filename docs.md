# DOCX Image Converter Tool Documentation

## Overview
This Python-based tool allows users to convert images embedded within DOCX files to a more optimized format (JPG), resize them, and set a specific DPI (Dots Per Inch). It can handle PNG, BMP, TIFF, and WebP images and convert them to JPG to reduce file size. It also updates the necessary references in the DOCX file, ensuring the images are properly linked after conversion.

## Features:
- **Image Conversion**: Converts supported image formats (PNG, BMP, TIFF, WebP) in DOCX files to JPG.
- **Image Resizing**: Resizes images to a target width based on a DPI setting.
- **Transparency Handling**: Skips images with transparency to avoid losing transparency information during conversion.
- **File Size Reduction**: The converted DOCX file generally has a smaller size due to image compression and conversion to JPG.
- **Rels Update**: Updates DOCX's internal XML references to point to the newly converted images.
  
## Usage Instructions

### 1. File Paths
The program accepts both **absolute** (full) and **relative** file paths when specifying the location of the DOCX file to be processed.

- **Absolute Path**: e.g., `C:/Documents/myfile.docx`
- **Relative Path**: e.g., `uploads/myfile.docx` (relative to the folder where the script is run)

### 2. Quality Settings
The image quality setting used during conversion is configured through the `settings.cfg` file. This value determines the compression quality of the JPG images created from the original images in the DOCX.

- **Quality Setting Range**: The value should be between **1** (lowest quality, highest compression) and **100** (highest quality, least compression).
  
In general, a quality setting of **75-85** offers a good balance between image clarity and file size reduction.

### 3. DPI Settings
The DPI (Dots Per Inch) value, also specified in the `settings.cfg` file, determines the resolution of the converted images. This setting impacts the size of the converted image, with higher DPI values resulting in larger and more detailed images.

- **Default DPI**: Typically set around **200 DPI** for high-quality prints.

### 4. Settings Configuration
The conversion settings (image quality and DPI) are stored in the `settings.cfg` file. The relevant section of the configuration file looks like this:

```ini
[ImageSettings]
quality = 85  # Quality out of 100
dpi = 200     # DPI setting for the images
```

You can modify these settings as needed. For example, reducing the `quality` value will decrease the file size further, at the cost of image clarity.

### 5. Running the Program
To run the program, execute the Python script and follow the on-screen prompts:

```bash
python convert_docx_images.py
```

When prompted, provide the path to the `.docx` file you want to convert. You can either use the full (absolute) path or the relative path from your current working directory.

Example input paths:
- Absolute path: `C:/Users/username/Documents/sample.docx`
- Relative path: `uploads/sample.docx`

### 6. Output
The converted `.docx` file will be saved in the `output` folder with a filename prefixed by `converted_`, e.g., `converted_sample.docx`.

### 7. File Size Reduction Display
Once the conversion is complete, the program displays the file size reduction achieved through the process. The display includes the original file size, the converted file size, and the percentage reduction.

## Key Functions

### 1. `allowed_file(filename)`
- Ensures the uploaded file has a `.docx` extension.

### 2. `has_transparency(image)`
- Checks if an image contains transparent pixels (applies only to images with an alpha channel).

### 3. `resize_image(image, target_width, dpi)`
- Resizes the image to a target width based on the given DPI.

### 4. `convert_docx_images(docx_filepath, output_filepath, quality, dpi)`
- Handles the extraction of images from the DOCX file, converts them to JPG, resizes them, and repackages the DOCX.

### 5. `update_rels_for_image_conversion(rels_filepath, original_image, converted_image)`
- Updates the references in DOCX `.rels` files to point to the newly converted images.

### 6. `display_file_size_reduction(original_filepath, converted_filepath)`
- Displays a comparison of the original and converted file sizes and the percentage space saved.

## Settings Configuration
Before running the program, you should verify or adjust the following settings in the `settings.cfg` file:

```ini
[ImageSettings]
quality = 85  # The quality of the converted images (out of 100)
dpi = 200     # The DPI (resolution) for the converted images
```

## Troubleshooting
- **Invalid DOCX File**: If you receive an error indicating the file is invalid or corrupted, ensure that the file is a proper DOCX file and not password protected.
- **Missing Media Directory**: If the media directory or `_rels` directory is missing from the DOCX, ensure the document contains embedded images.

## Conclusion
This tool is an efficient way to optimize DOCX files by converting embedded images to a more manageable format (JPG), resizing them, and reducing the overall file size while maintaining quality. By using the configuration options, users can control the balance between file size and image quality.