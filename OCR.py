import os
import win32com.client
import time

def create_output_folder_structure(top_folder, output_top_folder):
    """Create the output folder structure to mirror the input folder structure."""
    for root, dirs, files in os.walk(top_folder):
        for dir in dirs:
            output_dir_path = os.path.join(output_top_folder, os.path.relpath(os.path.join(root, dir), top_folder))
            os.makedirs(output_dir_path, exist_ok=True)

def process_pdf_files(top_folder, output_top_folder):
    """Iterate through each folder and process the PDF files."""
    # Initialize the Adobe Acrobat COM object
    acrobat_app = win32com.client.Dispatch("AcroExch.App")
    
    for root, dirs, files in os.walk(top_folder):
        for file in files:
            if file.endswith(".pdf"):
                input_file_path = os.path.join(root, file)
                relative_path = os.path.relpath(root, top_folder)
                output_dir_path = os.path.join(output_top_folder, relative_path)
                output_file_path = os.path.join(output_dir_path, file)

                try:
                    print(f"Processing: {input_file_path}")
                    pdf_document = win32com.client.Dispatch("AcroExch.PDDoc")
                    pdf_document.Open(input_file_path)

                    # Perform OCR using Acrobat
                    ocr_success = pdf_document.ApplyOCR()

                    if ocr_success:
                        # Save the processed PDF to the output directory
                        pdf_document.SaveAs(output_file_path, 1)  # 1 is for SaveAsCopy
                        print(f"Saved OCR PDF to: {output_file_path}")
                    else:
                        print(f"OCR failed for: {input_file_path}")

                    pdf_document.Close()

                except Exception as e:
                    print(f"Error processing {input_file_path}: {e}")

                # Give Acrobat a short break to prevent overwhelming it
                time.sleep(1)

    # Clean up the Acrobat application
    acrobat_app.Exit()

if __name__ == "__main__":
    top_folder = "C:\\path\\to\\your\\TOP\\folder"  # Update with your TOP folder path
    output_top_folder = "C:\\path\\to\\your\\output\\folder"  # Update with your output folder path

    # Ensure the output folder structure mirrors the input folder structure
    create_output_folder_structure(top_folder, output_top_folder)

    # Process the PDF files
    process_pdf_files(top_folder, output_top_folder)

    print("Processing completed.")
