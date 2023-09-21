import os
import re
from datetime import datetime

import pandas as pd
from PyPDF2 import PdfReader


def extract_expenses(pdf_folder_path):
    """Extract tables from PDF files in a folder using PyPDF2."""
    tabulated_data = []

    # Loop through each PDF file in the folder
    for pdf_filename in os.listdir(pdf_folder_path):
        if pdf_filename.endswith(".pdf") or pdf_filename.endswith(".PDF"):
            pdf_path = os.path.join(pdf_folder_path, pdf_filename)

            # Read the PDF
            with open(pdf_path, "rb") as pdf_file:
                pdf_reader = PdfReader(pdf_file)
                page = pdf_reader.pages[0]
                text = page.extract_text()

                # Extract and format the date from the 3rd line of the entire text
                raw_date_line = text.split("\n")[2].strip()
                raw_date_match = re.search(r"(\d{2}-\d{2}-\d{2})", raw_date_line)
                if raw_date_match:
                    raw_date = raw_date_match.group(1)
                    formatted_date = datetime.strptime(raw_date, "%d-%m-%y").strftime(
                        "%Y-%m-%d"
                    )
                else:
                    formatted_date = "Unknown Date"

                # Extract lines between "VENTE CAISSIER" and "CODE D'AUT"
                start = text.find("VENTE CAISSIER")
                end = text.find("CODE D'AUT")
                relevant_lines = text[start:end].split("\n")[1:-1]

                i = 0
                while i < len(relevant_lines):
                    line = relevant_lines[i]
                    if "<A>" in line:
                        item_code = line[:14].strip()
                        description = line[14:].split("<A>")[0].strip()
                        next_line_index = i + 1
                        if (
                            next_line_index < len(relevant_lines)
                            and "@" in relevant_lines[next_line_index]
                        ):
                            quantity, unit_price = relevant_lines[
                                next_line_index
                            ].split("@")
                            total = (
                                relevant_lines[next_line_index]
                                .split()[-1]
                                .replace(",", ".")
                            )
                            unit_price = unit_price.split()[0].replace(",", ".")
                            i += 1
                        else:
                            quantity = 1
                            unit_price = line.split("<A>")[-1].replace(",", ".")
                            total = unit_price
                        tabulated_data.append(
                            [
                                "Home Depot",
                                formatted_date,
                                pdf_filename,
                                item_code,
                                description,
                                quantity,
                                unit_price,
                                total,
                                "",
                                "",
                            ]
                        )
                    i += 1

                # Extract amounts for SOUS-TOTAL, TPS/TVH, TVP/TVQ, and TOTAL
                total_start = text.find("SOUS-TOTAL")
                total_end = text.find("CAD$")
                total_info = text[total_start:total_end].split("\n")

                sous_total = total_info[0].split()[-1].replace(",", ".")
                tps_tvh = total_info[1].split()[-1].replace(",", ".")
                tvp_tvq = total_info[2].split()[-1].replace(",", ".")
                total = total_info[3].split()[-1].replace(",", ".")

                # Add amounts to the tabulated data
                tabulated_data.append(
                    [
                        "Home Depot",
                        formatted_date,
                        pdf_filename,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "SOUS TOTAL",
                        sous_total,
                    ]
                )
                tabulated_data.append(
                    [
                        "Home Depot",
                        formatted_date,
                        pdf_filename,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "TPS/TVH",
                        tps_tvh,
                    ]
                )
                tabulated_data.append(
                    [
                        "Home Depot",
                        formatted_date,
                        pdf_filename,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "TVP/TVQ",
                        tvp_tvq,
                    ]
                )
                tabulated_data.append(
                    [
                        "Home Depot",
                        formatted_date,
                        pdf_filename,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "TOTAL",
                        total,
                    ]
                )

    # Create a DataFrame and save it to an Excel file
    data_table = pd.DataFrame(
        tabulated_data,
        columns=[
            "Store",
            "Date",
            "File Name",
            "Item Code",
            "Description",
            "Quantity",
            "Unit Price",
            "Total",
            "TextSum",
            "Sum",
        ],
    )
    data_table["Quantity"] = pd.to_numeric(data_table["Quantity"], errors="coerce")
    data_table["Unit Price"] = pd.to_numeric(data_table["Unit Price"], errors="coerce")
    data_table["Total"] = pd.to_numeric(
        data_table["Total"].str.replace("$", ""), errors="coerce"
    )
    data_table["Sum"] = pd.to_numeric(
        data_table["Sum"].str.replace("$", ""), errors="coerce"
    )

    return data_table


results = extract_expenses("receipts/HomeDepot")
results.to_excel("receipts/home-depot_data.xlsx", index=False)
print("Tabulated data has been saved to 'home-depot_data.xlsx'")
