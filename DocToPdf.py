import os
import win32com.client

# This script turns Microsoft Word files into PDF files
# It handles .docx and .doc files.
# If files are open or password-protected, it may skip or error out

# Folder path here
folder_path = r"C:\Users\CASPER\OneDrive\Masaüstü\DOCUMENTS-20250218"

# Launch Word application
word = win32com.client.Dispatch("Word.Application")
word.Visible = 0  # Set to 1 if you want to see Word open the files

# Loop through all Word documents
for file_name in os.listdir(folder_path):
    if file_name.endswith(".docx") or file_name.endswith(".doc"):
        full_path = os.path.join(folder_path, file_name)
        pdf_path = os.path.splitext(full_path)[0] + ".pdf"

        print(f"Converting: {file_name} -> {os.path.basename(pdf_path)}")

        try:
            document = word.Documents.Open(full_path)
            document.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
            document.Close()
        except Exception as e:
            print(f" Failed to convert {file_name}: {e}")

# Quit Word
word.Quit()

print("✅ All Word documents converted to PDF.")
