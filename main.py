import os
from tkinter import Tk, filedialog
import win32com.client
from pdf2docx import Converter
from pathlib import Path


def select_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename()
    return file_path


def select_destination_folder():
    root = Tk()
    root.withdraw()  # Hide the main window

    folder_path = filedialog.askdirectory(title="Select Destination Folder")
    return folder_path


def pdf_to_docx(pdf_path, docx_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        print(f"Conversion successful: {docx_path}")
    except Exception as e:
        print(f"Error converting PDF to DOCX: {e}")


def docx_to_pdf(docx_path, pdf_path):
    print(docx_path, "\n", pdf_path)
    docx_path = str(docx_path).replace("/", "\\")
    print(docx_path)
    pdf_path = str(pdf_path).replace("/", "\\")
    print(pdf_path)
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17)  # 17 represents PDF format
    doc.Close()
    word.Quit()
    print(f"Conversion successful: {pdf_path}")


def convert_file(source_path, destination_folder):
    source_file_name = os.path.basename(source_path)
    source_file_extension = os.path.splitext(source_path)[1].lower()

    destination_folder = Path(destination_folder)

    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    if source_file_extension == ".pdf":
        destination_extension = ".docx"
        conversion_function = pdf_to_docx
    elif source_file_extension in [".doc", ".docx"]:
        destination_extension = ".pdf"
        conversion_function = docx_to_pdf
    else:
        print("Unsupported file format")
        return

    destination_file_name = source_file_name.replace(source_file_extension, destination_extension)
    destination_path = destination_folder / destination_file_name

    conversion_function(source_path, destination_path)


def main():
    source_file = select_file()
    destination_folder = select_destination_folder()

    convert_file(source_file, destination_folder)


if __name__ == "__main__":
    main()
