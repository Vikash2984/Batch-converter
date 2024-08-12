import os
import win32com.client

def convert_pptx_to_pdf(source_folder, target_folder):
    # Initialize PowerPoint application
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    
    # Ensure the target folder exists
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # Iterate through all files in the source folder
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            if file.endswith('.pptx'):
                file_path = os.path.join(root, file)
                try:
                    # Open the PowerPoint file
                    presentation = powerpoint.Presentations.Open(file_path, ReadOnly=True)
                    
                    # Define the output PDF path
                    pdf_path = os.path.join(target_folder, file.replace('.pptx', '.pdf'))
                    
                    # Save as PDF
                    presentation.SaveAs(pdf_path, 32)  # 32 is the PDF format
                    
                    # Close the presentation
                    presentation.Close()
                    
                    print(f"Converted {file_path} to {pdf_path}")
                except Exception as e:
                    print(f"Failed to convert {file_path}: {e}")

    # Quit PowerPoint application
    powerpoint.Quit()

# Specify the source and target folder paths
source_folder = r'D:\vsCode\Python\Auto\Presentations'
target_folder = r'D:\vsCode\Python\Auto\PDFs'

# Convert all .pptx files in the source folder to PDF and save them to the target folder
convert_pptx_to_pdf(source_folder, target_folder)
