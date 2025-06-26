from pypdf import PdfReader
from datetime import datetime

def convert_pdf_date(pdf_date):
    if not pdf_date:
        return ""
    try:
        if pdf_date.startswith("D:"):
            pdf_date = pdf_date[2:]
        if "+" in pdf_date or "-" in pdf_date:
            sign = "+" if "+" in pdf_date else "-"
            main, tz = pdf_date.split(sign)
            tz = tz.replace("'", "")
            pdf_date = f"{main}{sign}{tz}"
            dt = datetime.strptime(pdf_date, "%Y%m%d%H%M%S%z")
        else:
            dt = datetime.strptime(pdf_date[:14], "%Y%m%d%H%M%S")
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return pdf_date

def extract_metadata(folder_path):
    import os
    metadata_list = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                try:
                    reader = PdfReader(file_path)
                    metadata = reader.metadata or {}
                    metadata_list.append({
                        "File Name": file,
                        "Full Path": file_path,
                        "Title": metadata.get("/Title", ""),
                        "Author": metadata.get("/Author", ""),
                        "Subject": metadata.get("/Subject", ""),
                        "Creator": metadata.get("/Creator", ""),
                        "Producer": metadata.get("/Producer", ""),
                        "Creation Date": convert_pdf_date(metadata.get("/CreationDate", "")),
                        "Modification Date": convert_pdf_date(metadata.get("/ModDate", ""))
                    })
                except Exception as e:
                    metadata_list.append({
                        "File Name": file,
                        "Full Path": file_path,
                        "Title": "",
                        "Author": "",
                        "Subject": "",
                        "Creator": "",
                        "Producer": "",
                        "Creation Date": "",
                        "Modification Date": "",
                        "Error": str(e)
                    })
    return metadata_list