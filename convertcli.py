import zipfile
import os
import shutil
from PIL import Image
import logging
import configparser

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Set the upload folder and output folder paths
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
ALLOWED_EXTENSIONS = {'docx'}
ALLOWED_IMAGE_EXTENSIONS = {'.png', '.bmp', '.tiff', '.webp'}

# Load settings from the config file
config = configparser.ConfigParser()
config.read('settings.cfg')

# Read quality and DPI from the configuration file
IMAGE_QUALITY = config.getint('ImageSettings', 'quality')
IMAGE_DPI = config.getint('ImageSettings', 'dpi')

def allowed_file(filename):
    """Check if the uploaded file is a .docx file."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def has_transparency(image):
    """Check if an image has any transparent pixels."""
    if image.mode != 'RGBA':
        image = image.convert('RGBA')
    
    alpha_channel = image.getchannel('A')
    return alpha_channel.getextrema()[0] < 255

def resize_image(image, target_width, dpi):
    """Resize image to target width and set the DPI."""
    target_height = int((target_width / image.width) * image.height)
    resized_image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)
    return resized_image

def update_rels_for_image_conversion(rels_filepath, original_image, converted_image):
    """Update the image extension in the rels XML from .png, .webp, etc. to .jpg."""
    with open(rels_filepath, 'r', encoding='utf-8') as f:
        rels_content = f.read()
    updated_rels_content = rels_content.replace(original_image, converted_image)
    with open(rels_filepath, 'w', encoding='utf-8') as f:
        f.write(updated_rels_content)

def update_all_rels_files(rels_dir, original_image, converted_image):
    """Update all .rels files in the word/_rels directory with the new image reference."""
    for filename in os.listdir(rels_dir):
        if filename.endswith('.rels'):
            rels_filepath = os.path.join(rels_dir, filename)
            update_rels_for_image_conversion(rels_filepath, original_image, converted_image)

def convert_docx_images(docx_filepath, output_filepath, quality, dpi):
    """Convert PNG, BMP, TIFF, or WebP images in a docx file to JPG and resize."""
    # Create a temporary directory to extract the docx contents
    temp_dir = 'temp_docx'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    try:
        # Unzip the .docx file into the temp directory
        with zipfile.ZipFile(docx_filepath, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
    except zipfile.BadZipFile:
        logging.error(f"Error: The file '{docx_filepath}' is not a valid DOCX or is corrupted.")
        return False

    # Path to the relationships directory
    rels_dir = os.path.join(temp_dir, 'word/_rels')
    
    if not os.path.exists(rels_dir):
        logging.error(f"Error: The '_rels' directory does not exist in the DOCX structure.")
        shutil.rmtree(temp_dir)
        return False
    
    # Process images in the word/media directory
    media_dir = os.path.join(temp_dir, 'word/media')
    if not os.path.exists(media_dir):
        logging.error(f"Error: The 'media' directory does not exist in the DOCX structure.")
        shutil.rmtree(temp_dir)
        return False

    image_count = 0
    for filename in os.listdir(media_dir):
        file_extension = os.path.splitext(filename)[1].lower()
        if file_extension in ALLOWED_IMAGE_EXTENSIONS:
            filepath = os.path.join(media_dir, filename)
            try:
                with Image.open(filepath) as image:
                    if not has_transparency(image):
                        # Resize and convert to JPG
                        target_width = 6*dpi  # 6 inches * 200 PPI
                        resized_image = resize_image(image.convert('RGB'), target_width, dpi=dpi)
                        
                        # Save the new image as JPG
                        new_filename = filename.rsplit('.', 1)[0] + '.jpg'
                        new_filepath = os.path.join(media_dir, new_filename)
                        resized_image.save(new_filepath, format='JPEG', quality=quality, dpi=(dpi, dpi))

                        # Remove the old image
                        os.remove(filepath)

                        # Update all rels files to point to the new JPG
                        update_all_rels_files(rels_dir, f'media/{filename}', f'media/{new_filename}')
                        
                        image_count += 1
            except Exception as e:
                logging.error(f"Error processing image {filename}: {e}")
    
    if image_count == 0:
        logging.info("No valid images were found to convert.")
    else:
        logging.info(f"Converted {image_count} images successfully.")

    # Repack the folder into a .docx file
    with zipfile.ZipFile(output_filepath, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, temp_dir)
                docx_zip.write(file_path, arcname)

    # Clean up the temporary directory
    shutil.rmtree(temp_dir)
    
    return True

def display_file_size_reduction(original_filepath, converted_filepath):
    """Display the size reduction between original and converted files."""
    original_size = os.path.getsize(original_filepath)
    converted_size = os.path.getsize(converted_filepath)
    
    size_difference = original_size - converted_size
    percentage_reduction = (size_difference / original_size) * 100
    
    print(f"Original file size: {original_size / 1024:.2f} KB")
    print(f"Converted file size: {converted_size / 1024:.2f} KB")
    print(f"Space saved: {size_difference / 1024:.2f} KB ({percentage_reduction:.2f}% reduction)")

def main():
    """Main loop for handling file conversion via the CLI."""
    while True:
        # Prompt user for input file path
        input_path = input("Enter the path to the .docx file (or type 'exit' to quit): ")
        if input_path.lower() == 'exit':
            print("Exiting the program. Goodbye!")
            break
        
        if not os.path.exists(input_path):
            print(f"Error: The file '{input_path}' does not exist. Please try again.")
            continue

        if not allowed_file(input_path):
            print(f"Error: The file '{input_path}' is not a valid .docx file. Please upload a .docx file.")
            continue

        # Output file path
        output_filename = f"converted_{os.path.basename(input_path)}"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        # Convert the docx images with the quality and DPI from the config file
        print("Converting images in the document...")
        success = convert_docx_images(input_path, output_path, quality=IMAGE_QUALITY, dpi=IMAGE_DPI)
        
        if success:
            print(f"File has been successfully converted and saved as: {output_path}")
            # Display size reduction information
            display_file_size_reduction(input_path, output_path)
        else:
            print("File conversion failed.")
        
        # Ask if the user wants to convert another file
        again = input("Would you like to convert another file? (y/n): ").lower()
        if again != 'y':
            print("Exiting the program. Goodbye!")
            break

if __name__ == "__main__":
    main()
