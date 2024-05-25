import os
import glob
import subprocess

# Function to convert files
def convert_ppt_file(input_path, output_folder):
    try:
        # Get the current filename and extension
        current_filename, current_extension = os.path.splitext(os.path.basename(input_path))
        input_extension = current_extension.lower()

        # Set the output filename and path
        output_filename = f"{current_filename}.pdf"
        output_path = os.path.join(output_folder, output_filename)

        cmd = ["libreoffice", "--headless", "--convert-to", 'pdf', "--outdir", output_folder, input_path]
        status = subprocess.call(cmd)
        if status == 0:
            return output_path
        else:
            return None
    except:
        return None

def convert_doc_to_docx(doc_path, output_folder):
    try:
        # Ensure the path is absolute
        current_filename, current_extension = os.path.splitext(os.path.basename(doc_path))
        input_extension = current_extension.lower()

        # Set the output filename and path
        output_filename = f"{current_filename}.docx"
        output_path = os.path.join(output_folder, output_filename)

        # Run the LibreOffice conversion command
        cmd = ["libreoffice", "--headless", "--convert-to", "docx", "--outdir", output_folder, doc_path]
        status = subprocess.call(cmd)
        
        if status == 0:
            return output_path
        else:
            return None
    except:
        return None

def convert_xls_to_xlsx(xls_path, output_folder):
    try:
        # Ensure the path is absolute
        current_filename, current_extension = os.path.splitext(os.path.basename(xls_path))
        input_extension = current_extension.lower()

        # Set the output filename and path
        output_filename = f"{current_filename}.xlsx"
        output_path = os.path.join(output_folder, output_filename)

        # Run the LibreOffice conversion command
        cmd = ["libreoffice", "--headless", "--convert-to", "xlsx", "--outdir", output_folder, xls_path]
        status = subprocess.call(cmd)
        
        if status == 0:
            return output_path
        else:
            return None
    except:
        return None

def convert_ppt_to_pptx(ppt_path, output_folder):
    try:
        # Ensure the path is absolute
        current_filename, current_extension = os.path.splitext(os.path.basename(ppt_path))
        input_extension = current_extension.lower()

        # Set the output filename and path
        output_filename = f"{current_filename}.pptx"
        output_path = os.path.join(output_folder, output_filename)

        # Run the LibreOffice conversion command
        cmd = ["libreoffice", "--headless", "--convert-to", "pptx", "--outdir", output_folder, ppt_path]
        status = subprocess.call(cmd)
        
        if status == 0:
            return output_path
        else:
            return None
    except:
        return None

