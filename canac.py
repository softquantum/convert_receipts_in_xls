import os
import re
from datetime import datetime

import pandas as pd
import pdfplumber


def extract_expenses(pdf_folder_path):
    """Extract tables from PDF files in a folder using pdfplumber."""
    columns = [
        "Store",
        "Date",
        "Filename",
        "Article",
        "Description",
        "Quantité",
        "UdM",
        "Prix Unité",
        "Total",
        "TextSum",
        "Sum",
    ]
    data_table = pd.DataFrame(columns=columns)

    # Loop through each PDF file in the folder
    for pdf_file_name in os.listdir(pdf_folder_path):
        if pdf_file_name.endswith(".PDF") or pdf_file_name.endswith(".pdf"):
            pdf_file = os.path.join(pdf_folder_path, pdf_file_name)

            # Open the PDF file with pdfplumber
            with pdfplumber.open(pdf_file) as pdf:
                # Loop through each page in the PDF
                for page in pdf.pages:
                    # Extract the text
                    text = page.extract_text()

                    # Extract and format the date from the 4th line of the entire text
                    raw_date_line = text.split("\n")[4].strip()
                    raw_date_match = re.search(r"(\d{4}/\d{2}/\d{2})", raw_date_line)
                    if raw_date_match:
                        raw_date = raw_date_match.group(1)
                        formatted_date = datetime.strptime(
                            raw_date, "%Y/%m/%d"
                        ).strftime("%Y-%m-%d")
                    else:
                        formatted_date = "Unknown Date"

                    # Extract lines between "Article Description Quantité UdM Prix Unité Total" and "Mastercard"
                    start = text.find(
                        "Article Description Quantité UdM Prix Unité Total"
                    )
                    end = text.find("Mastercard")
                    relevant_lines = text[start:end].split("\n")[1:-1]

                    # Split the text by lines and iterate through each line
                    rows = []
                    for line in relevant_lines:
                        # Split the line by whitespace to get the individual elements
                        elements = line.split()
                        if len(elements) >= 6:
                            row = {
                                "Store": "Canac",
                                "Date": formatted_date,
                                "Filename": pdf_file_name,
                                "Article": elements[0],
                                "Description": " ".join(elements[1:-4]),
                                "Quantité": elements[-4],
                                "UdM": elements[-3],
                                "Prix Unité": elements[-2],
                                "Total": elements[-1],
                                "TextSum": "",
                                "Sum": "",
                            }
                            rows.append(row)

                        # Extract amounts for SOUS-TOTAL, TPS/TVH, TVP/TVQ, and TOTAL
                        elif len(elements) <= 4:
                            row = {
                                "Store": "Canac",
                                "Date": formatted_date,
                                "Filename": pdf_file_name,
                                "Article": "",
                                "Description": "",
                                "Quantité": "",
                                "UdM": "",
                                "Prix Unité": "",
                                "Total": "",
                                "TextSum": " ".join(elements[:-1]),
                                "Sum": elements[-1],
                            }
                            rows.append(row)

                    # Create a DataFrame from the rows
                    data = pd.DataFrame(rows, columns=columns)

                    # Concatenate the DataFrame to the main one
                    data_table = pd.concat([data_table, data], ignore_index=True)

    data_table["Quantité"] = pd.to_numeric(data_table["Quantité"], errors="coerce")
    data_table["Prix Unité"] = pd.to_numeric(data_table["Prix Unité"], errors="coerce")
    data_table["Total"] = pd.to_numeric(data_table["Total"], errors="coerce")
    data_table["Sum"] = pd.to_numeric(data_table["Sum"], errors="coerce")

    return data_table


results = extract_expenses("receipts/Canac")
results.to_excel("receipts/canac_data.xlsx", index=False)
print("Tabulated data has been saved to 'canac_data.xlsx'")
