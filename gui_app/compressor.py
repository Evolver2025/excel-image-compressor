import os
import zipfile
import shutil
from PIL import Image
import argparse
import uuid
import time
import io

def compress_excel_images(file_path, compression_level=20, png_colors=64, convert_png_to_jpeg=True, logger=print, progress_callback=None):
    """
    Compresses images within an Excel file in memory.

    :param file_path: Path to the Excel file.
    :param compression_level: Compression quality (0-100).
    :param progress_callback: A function to call with progress updates (0-100).
    """
    if not os.path.exists(file_path):
        logger(f"Error: File not found at {file_path}")
        return

    file_path = os.path.abspath(file_path)
    base_dir = os.path.dirname(file_path)
    file_name_no_ext = os.path.splitext(os.path.basename(file_path))[0]
    output_filename = f"{file_name_no_ext}_compressed.xlsx"
    output_path = os.path.join(base_dir, output_filename)

    # In-memory storage for zip contents
    zip_in_memory = {}

    try:
        # 1. Read the original Excel (zip) file into memory
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            for item in zip_ref.infolist():
                zip_in_memory[item.filename] = zip_ref.read(item.filename)

        # 2. Identify and compress images
        media_path_prefix = "xl/media/"
        image_files = [name for name in zip_in_memory if name.startswith(media_path_prefix)]
        
        if not image_files:
            logger("No images found in xl/media/")
            # If no images, we can just copy the file, or simply exit.
            # To be safe and provide a result, let's create a "compressed" copy.
            shutil.copy(file_path, output_path)
            logger(f"\nNo images to compress. Copied to: {output_path}")
            return

        num_images = len(image_files)
        total_original_size = 0
        total_compressed_size = 0

        for i, image_name in enumerate(image_files):
            original_data = zip_in_memory[image_name]
            original_size = len(original_data)
            total_original_size += original_size
            
            try:
                with Image.open(io.BytesIO(original_data)) as img:
                    original_format = img.format
                    output_buffer = io.BytesIO()
                    save_options = {}

                    if img.format == 'PNG' and convert_png_to_jpeg:
                        logger(f"Converting PNG to JPEG: {os.path.basename(image_name)}")
                        if img.mode in ('RGBA', 'LA'):
                            background = Image.new("RGB", img.size, (255, 255, 255))
                            background.paste(img, mask=img.split()[-1])
                            img = background
                        else:
                            img = img.convert('RGB')
                        save_options = {'quality': compression_level, 'optimize': True}
                        img.save(output_buffer, format='JPEG', **save_options)

                    elif img.format == 'JPEG':
                        save_options = {'quality': compression_level, 'optimize': True}
                        img.save(output_buffer, format=original_format, **save_options)
                    
                    elif img.format == 'PNG':
                        try:
                            img = img.quantize(colors=png_colors, method=Image.Quantize.LIBIMAGEQUANT)
                            save_options = {'optimize': True}
                            img.save(output_buffer, format=original_format, **save_options)
                        except Exception:
                            save_options = {'compress_level': 9, 'optimize': True}
                            img.save(output_buffer, format=original_format, **save_options)
                    else:
                        save_options = {'optimize': True}
                        img.save(output_buffer, format=original_format, **save_options)

                    compressed_data = output_buffer.getvalue()
                    compressed_size = len(compressed_data)
                    total_compressed_size += compressed_size
                    
                    # Update the in-memory zip with the compressed image
                    zip_in_memory[image_name] = compressed_data
                    
                    logger(f"Compressed {os.path.basename(image_name)}: {original_size / 1024:.2f} KB -> {compressed_size / 1024:.2f} KB")

            except Exception as e:
                logger(f"Could not compress {os.path.basename(image_name)}: {e}")
                # Keep original data if compression fails
                total_compressed_size += original_size

            if progress_callback:
                progress = ((i + 1) / num_images) * 100
                progress_callback(progress)

        if total_original_size > 0:
            reduction_percent = (1 - total_compressed_size / total_original_size) * 100
            logger(f"Total image size reduction: {total_original_size / 1024:.2f} KB -> {total_compressed_size / 1024:.2f} KB ({reduction_percent:.2f}%)")

        # 3. Write the new Excel file from memory
        if os.path.exists(output_path):
            os.remove(output_path)
            
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for filename, data in zip_in_memory.items():
                zip_out.writestr(filename, data)

        logger(f"\nSuccessfully created compressed file: {output_path}")

    except Exception as e:
        logger(f"An unexpected error occurred: {e}")
