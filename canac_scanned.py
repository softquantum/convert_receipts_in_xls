import os
import re
from datetime import datetime

import pandas as pd
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

TPS_PERCENTAGE = 0.05
TVQ_PERCENTAGE = 0.09975


def get_element(a_list, index):
    """Return the element at the given index if it exists, otherwise return the default value."""
    return a_list[index] if index < len(a_list) else None


def extract_numeric_value(str):
    """Extract the first numeric value from a string."""
    if str:
        matches = re.findall(r"\d+[\s]*[,\.]?[\s]*\d*", str)
        if matches:
            number = float(matches[0].replace(",", ".").replace(" ", ""))
            return int(number) if number == int(number) else number
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
    image = image.convert("L")
    image = image.filter(ImageFilter.MedianFilter(size=3))
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)
    enhancer = ImageEnhance.Sharpness(image)
    image = enhancer.enhance(1.5)
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
    data_table = pd.DataFrame(columns=columns)

    for file_name in os.listdir(file_folder_path):
        formatted_date = "Unknown Date"
        sous_total = tps = tvq = grand_total = None

        if file_name.endswith((".jpg", ".jpeg")):
            print(f"Extracting data from {file_name}...")
            image = prepare_image_for_ocr(os.path.join(file_folder_path, file_name))
            text = pytesseract.image_to_string(image, config="--oem 3 --psm 6")
            relevant_lines = text.split("\n")[1:-1] if text.split("\n")[1:-1] else []

            for line in text.split("\n"):
                raw_date_match = re.search(r"(\d{2}-\d{2}-\d{2})", line)
                if raw_date_match:
                    raw_date = raw_date_match.group(1)
                    try:
                        formatted_date = datetime.strptime(
                            raw_date, "%m-%d-%y"
                        ).strftime("%Y-%m-%d")
                    except ValueError:
                        try:
                            formatted_date = datetime.strptime(
                                raw_date, "%d-%m-%y"
                            ).strftime("%Y-%m-%d")
                        except ValueError:
                            formatted_date = raw_date
                    break

            i = 0
            c = d = 0
            item_code = description = quantity = unit_price = total = None
            sous_total = tps = tvq = grand_total = None

            while i < len(relevant_lines):
                line = relevant_lines[i].strip()

                if not line:
                    i += 1
                    continue

                if "#produit" in line or any(
                    sub in line
                    for sub in [
                        "duit",
                        "oduit",
                        "roduit",
                        "raduit",
                        "prod",
                        "oroduit",
                        "foraduit",
                        "odult",
                        "rodult",
                    ]
                ):
                    item_code = line.strip()
                    c = d = 1
                    i += 1
                    continue

                elements = line.split()
                if len(elements) == 1:
                    i += 1
                    continue

                if c == 1:
                    elements = line.split()
                    total = extract_numeric_value(elements[-1] if elements else "0")
                    description = (
                        " ".join(elements[:-1]) if total is not None else line.strip()
                    )
                    c = 0
                    i += 1
                    continue

                if "x" in line and d == 1:
                    elements = line.split("x")
                    if len(elements) == 2:
                        quantity_str, unit_price_str = elements
                    else:
                        quantity_str = elements[0]
                        unit_price_str = " ".join(elements[1:])
                    unit_price = extract_numeric_value(unit_price_str)
                    quantity = (
                        extract_numeric_value(quantity_str)
                        if quantity_str not in ["|", "l"]
                        else 1
                    )
                    if quantity is None:
                        quantity = 1
                    d = 0

                try:
                    if total is None:
                        total = quantity * unit_price
                except TypeError:
                    total = 0

                if item_code and description and quantity and unit_price and total:
                    tabulated_data.append(
                        [
                            "Canac",
                            formatted_date,
                            file_name,
                            item_code,
                            description,
                            quantity,
                            unit_price,
                            total,
                            "",
                            "",
                        ]
                    )
                    item_code = description = quantity = unit_price = total = None

                # Extract
                if "SOUS-TOTAL" in line:
                    sous_total = extract_numeric_value(line)
                elif "TPS" in line:
                    tps = extract_numeric_value(line)
                elif "TVQ" in line:
                    tvq = extract_numeric_value(line)
                elif "TOTAL" in line:
                    grand_total = extract_numeric_value(line)

                i += 1

        # If sous_total is missing but grand_total is present
        if not sous_total and grand_total:
            sous_total = grand_total / (TPS_PERCENTAGE + TVQ_PERCENTAGE + 1)

        # If tps is missing but sous_total or total are present
        if not tps or sous_total or grand_total:
            if grand_total:
                tps = (
                    grand_total / (TPS_PERCENTAGE + TVQ_PERCENTAGE + 1) * TPS_PERCENTAGE
                )
            elif sous_total:
                tps = sous_total * TPS_PERCENTAGE

        # If tvq is missing but sous_total or total are present
        if not tvq or sous_total or grand_total:
            if grand_total:
                tvq = (
                    grand_total / (TPS_PERCENTAGE + TVQ_PERCENTAGE + 1) * TVQ_PERCENTAGE
                )
            elif sous_total:
                tvq = sous_total * TVQ_PERCENTAGE

        # If grand_total is missing but sous_total is present
        if not grand_total and sous_total:
            grand_total = sous_total * (TPS_PERCENTAGE + TVQ_PERCENTAGE + 1)

        sous_total = round(sous_total, 2) if sous_total else None
        tps = round(tps, 2) if tps else None
        tvq = round(tvq, 2) if tvq else None
        grand_total = round(grand_total, 2) if grand_total else None

        values_dict = {
            "SOUS TOTAL": sous_total,
            "TPS": tps,
            "TVQ": tvq,
            "TOTAL": grand_total,
        }

        # Loop through the dictionary and append each item to the data table
        print("Extracted values:", values_dict)
        for text, value in values_dict.items():
            tabulated_data.append(
                [
                    "Canac",
                    formatted_date,
                    file_name,
                    "",
                    "",
                    "",
                    "",
                    "",
                    text,
                    value,
                ]
            )

    data_table = pd.DataFrame(tabulated_data, columns=columns)
    return data_table


results = extract_expenses("receipts/Canac")
results.to_excel("receipts/canac_data_scanned.xlsx", index=False)
print("Tabulated data has been saved to 'canac_data_scanned.xlsx'")
print(results)
