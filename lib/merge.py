from PyPDF2 import PdfFileMerger
import os

def merge_files(files, filename):
    # Merge files into single report
    merger = PdfFileMerger()
    for pdf in files:
        merger.append(pdf)
    merger.write(f"reports/{filename}.pdf")
    merger.close