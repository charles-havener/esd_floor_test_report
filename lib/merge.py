from PyPDF2 import PdfFileMerger
import os

def merge_files(files):
    
    # Merge files into single report
    merger = PdfFileMerger()
    for pdf in files:
        merger.append(pdf)
    merger.write("reports/report.pdf")
    merger.close

