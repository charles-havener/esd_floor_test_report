from PyPDF2 import PdfFileMerger
import os

def merge_files(files):
    
    # Merge files into single report
    merger = PdfFileMerger()
    for pdf in files:
        merger.append(pdf)
    merger.write("report.pdf")
    merger.close

    # Remove pdfs, only keeping final report
    for f in files:
        os.remove(f)

