import os
import win32com.client

#This script turns power point files into pdf files
#It handles .pptx and .ppt files.
#If files are open or password-protected, it may skip or error out

# Folder path here
folder_path = r"C:\Users\CASPER\OneDrive\Masaüstü\PRESENTATIONS-20250218"

#Launch PowerPoint
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = 1

# Loop through all .pptx files
for file_name in os.listdir(folder_path):
    if file_name.endswith(".pptx") or file_name.endswith(".ppt"):
        full_path = os.path.join(folder_path, file_name)
        pdf_path = os.path.splitext(full_path)[0] + ".pdf"

        print(f"Converting: {file_name} -> {os.path.basename(pdf_path)}")

        presentation = powerpoint.Presentations.Open(full_path, WithWindow=False)
        presentation.SaveAs(pdf_path, 32)  # 32 = PDF format
        presentation.Close()

# Quit PowerPoint
powerpoint.Quit()

print("✅ All presentations converted to PDF.")
