from PyPDF2 import PdfFileReader
import os
import regex
from pandas import DataFrame
import datetime

# Avläs PDF'er från path
p = r"C:\Users\SEOHKR\OneDrive - Sweco AB\Desktop\Python\Projekt\SFA_PDF_Avläsning\underlag"


def check_valid_path(path=None):
    """Läs av alla pdf i angiven sökväg"""

    src_path = None
    if path is None:
        src_path = input(r"Välj sökväg för att läsa pdf")
    else:
        src_path = path

    return src_path


def read_from_path(path):
    """Läser alla PDF som finns i dir_path"""
    file_data = {}

    for pdf in os.listdir(path):
        if pdf.endswith(".pdf"):
            pdf_with_path = os.path.join(path, pdf)
            # relativ sökväg till pdf, 2 mappar upp
            pdf_dir = "\\".join(["\\".join((path.split("\\")[-2:])), pdf])
            # läs pdf
            pdf_reader = PdfFileReader(pdf_with_path)
            pages = pdf_reader.pages
            try:
                for index, page in enumerate(pages):
                    page_in_file = {}
                    text = page.extract_text(index)
                    text_to_list = text.split("\n")

                    # filtrera text med REGEX
                    rgx_pattern = r"(TOTAL VIKT kg) (\d+)"
                    for text in text_to_list:
                        match = regex.search(rgx_pattern, text)
                        if match is not None:
                            # vikt i kg
                            total_vikt_kg = int(match.group(2))
                            page_in_file["FIL"] = pdf
                            page_in_file[f"Sida {index +1}"] = index + 1
                            page_in_file[f"VIKT"] = total_vikt_kg

                # lägg till i file_data
                file_data[pdf_dir] = page_in_file
            except Exception as e:
                return e

    return file_data


def write_excel_summary(data: dict):
    """Turn dictionary to Pandas PD, write to excel"""
    df = DataFrame.from_dict(
        data=data).T

    # Förbered excel
    output = r"C:\Users\SEOHKR\OneDrive - Sweco AB\Desktop\Python\Projekt\SFA_PDF_Avläsning\output"
    fil = f"SFA_Avläst_PDF_{datetime.date.today()}.xlsx"
    output_path = os.path.join(output, fil)
    # Skriv excel
    df.to_excel(excel_writer=output_path, sheet_name="total vikt",
                header=["FIL", "SIDA", "VIKT - KG"])


if __name__ == "__main__":
    valid_path = check_valid_path(p)
    data = read_from_path(valid_path)
    to_excel = write_excel_summary(data)
