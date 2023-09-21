import os
import re
from datetime import datetime

import pandas as pd
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter, ImageOps


def get_element(a_list, index):
    """Return the element at the given index if it exists, otherwise return the default value."""
    return a_list[index] if index < len(a_list) else None


def extract_numeric_value(str):
    """Extract the first numeric value from a string."""
    if str:
        # Extract numeric values (including commas and periods)
        matches = re.findall(r"\d+[\s]*[,\.]?[\s]*\d*", str)
        if matches:
            # Replace commas with periods for decimal numbers and remove spaces
            return float(matches[0].replace(",", ".").replace(" ", ""))
        else:
            return None
    else:
        return None


def correct_image_orientation(image_path):
    """Correct the orientation of an image based on its EXIF data."""
    image = Image.open(image_path)
    try:
        exif_data = image._getexif()
        orientation = exif_data.get(274)
        if orientation == 3:
            image = image.rotate(180, expand=True)
        elif orientation == 6:
            image = image.rotate(-90, expand=True)
        elif orientation == 8:
            image = image.rotate(90, expand=True)
    except (AttributeError, KeyError, IndexError):
        pass
    return image


def prepare_image_for_ocr(image_path):
    """Prepare an image for OCR by converting it to grayscale and enhancing contrast."""
    image = correct_image_orientation(image_path)

    # Convert to grayscale
    image = image.convert("L")

    # Apply a median filter for noise reduction (optional)
    image = image.filter(ImageFilter.MedianFilter(size=3))

    # Enhance contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(1.3)

    # Enhance brightness
    # enhancer = ImageEnhance.Brightness(image)
    # image = enhancer.enhance(1.2)

    # Enhance sharpness
    enhancer = ImageEnhance.Sharpness(image)
    image = enhancer.enhance(1.5)

    # Resize image
    # image = image.resize(
    #     (int(image.width * 1.5), int(image.height * 1.5)), Image.LANCZOS
    # )

    # Invert image (depending on the background)
    # image = ImageOps.invert(image)

    return image


def extract_expenses(file_folder_path):
    """Extract tables from scanned receipts in a folder using pytesseract."""
    tabulated_data = []
    columns = [
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
    ]
    data_table = pd.DataFrame(tabulated_data, columns=columns)

    # Loop through each Image file in the folder
    for file_name in os.listdir(file_folder_path):
        if file_name.endswith(".jpg") or file_name.endswith(".jpeg"):
            image = prepare_image_for_ocr(os.path.join(file_folder_path, file_name))
            text = pytesseract.image_to_string(image)
            print("texte: ", text)

            # Extract lines between "VENTE CAISSIER" and "CODE D'AUT"
            start = text.find("VENTE C")
            end = text.find("CODE D")
            relevant_lines = text[start:end].split("\n")[1:-1]

            formatted_date = "Unknown Date"

            # Extract and format the date
            for line in text.split("\n"):
                raw_date_match = re.search(r"(\d{2}-\d{2}-\d{2})", line)
                if raw_date_match:
                    raw_date = raw_date_match.group(1)
                    formatted_date = datetime.strptime(raw_date, "%d-%m-%y").strftime(
                        "%Y-%m-%d"
                    )
                    break  # Stop the loop once the first date is found

            print("date: ", formatted_date)

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
                        quantity, unit_price = relevant_lines[next_line_index].split(
                            "@"
                        )
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
                            image,
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
            total_info = [x for x in total_info if x]

            sous_total = extract_numeric_value(get_element(total_info, 0))
            tps = extract_numeric_value(get_element(total_info, 1))
            tvq = extract_numeric_value(get_element(total_info, 2))

            if sous_total is not None:
                calculated_tps = round(sous_total * 0.05, 2)
                calculated_tvq = round(sous_total * 0.09975, 2)
                tps = calculated_tps if tps is None or tps != calculated_tps else tps
                tvq = calculated_tvq if tvq is None or tvq != calculated_tvq else tvq
                total = sous_total + tps + tvq
            else:
                total = extract_numeric_value(get_element(total_info, 3))

            print(sous_total, tps, tvq, total)

            # Add amounts to the tabulated data
            tabulated_data.append(
                [
                    "Home Depot",
                    formatted_date,
                    image,
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
                    image,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "TPS/TVH",
                    tps,
                ]
            )
            tabulated_data.append(
                [
                    "Home Depot",
                    formatted_date,
                    image,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "TVP/TVQ",
                    tvq,
                ]
            )
            tabulated_data.append(
                [
                    "Home Depot",
                    formatted_date,
                    image,
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
        data_table = pd.DataFrame(tabulated_data, columns=columns)
        # data_table["Quantity"] = pd.to_numeric(data_table["Quantity"], errors="coerce")
        # data_table["Unit Price"] = pd.to_numeric(
        #     data_table["Unit Price"], errors="coerce"
        # )
        # data_table["Total"] = pd.to_numeric(
        #     data_table["Total"].str.replace("$", ""), errors="coerce"
        # )
        # data_table["Sum"] = pd.to_numeric(
        #     data_table["Sum"].str.replace("$", ""), errors="coerce"
        # )

    return data_table


results = extract_expenses("receipts/HomeDepot")
# results.to_excel("receipts/homedepot_data_scanned.xlsx", index=False)
print("Tabulated data has been saved to 'homedepot_data_scanned.xlsx'")
