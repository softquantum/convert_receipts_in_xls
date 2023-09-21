import os
import re
from datetime import datetime

import cv2
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
            number = float(matches[0].replace(",", ".").replace(" ", ""))

            # Return an integer if the number is an integer, otherwise return a float
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
    # image = Image.open(image_path)

    # Convert to grayscale
    image = image.convert("L")

    # Apply a median filter for noise reduction (optional)
    image = image.filter(ImageFilter.MedianFilter(size=3))

    # Enhance contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)

    # Enhance brightness
    # enhancer = ImageEnhance.Brightness(image)
    # image = enhancer.enhance(1.1)

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
    data_table = pd.DataFrame(columns=columns)

    for file_name in os.listdir(file_folder_path):
        print("file_name", file_name)
        if file_name.endswith((".jpg", ".jpeg")):
            image = prepare_image_for_ocr(os.path.join(file_folder_path, file_name))

            # Using pytesseract
            text = pytesseract.image_to_string(image, config="--oem 3 --psm 6")

            print("text", text)
            relevant_lines = text.split("\n")[1:-1] if text.split("\n")[1:-1] else []
            formatted_date = "Unknown Date"

            for line in text.split("\n"):
                raw_date_match = re.search(r"(\d{2}-\d{2}-\d{2})", line)
                if raw_date_match:
                    raw_date = raw_date_match.group(1)
                    formatted_date = datetime.strptime(raw_date, "%m-%d-%y").strftime(
                        "%Y-%m-%d"
                    )
                    break

            i = 0
            c = d = 0
            item_code = description = quantity = unit_price = total = None

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

                # Skip lines with only one element
                elements = line.split()
                if len(elements) == 1:
                    print("Skipping line:", line)
                    i += 1
                    continue

                if c == 1:
                    line_elements = line.split()
                    total = extract_numeric_value(
                        line_elements[-1] if line_elements else "0"
                    )
                    description = (
                        " ".join(line_elements[:-1])
                        if total is not None
                        else line.strip()
                    )
                    c = 0
                    i += 1
                    continue

                if "x" in line and d == 1:
                    quantity_str, unit_price_str = line.split("x")
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

                if item_code is not None and description is not None:
                    if quantity is None:
                        try:
                            quantity = total / unit_price
                        except TypeError:
                            quantity = 1
                    if unit_price is None:
                        try:
                            unit_price = total / quantity
                        except TypeError:
                            unit_price = "N/A"
                    if total is None or total == 0:
                        try:
                            total = quantity * unit_price
                        except TypeError:
                            total = "N/A"

                    if quantity and unit_price:
                        total = (
                            quantity * unit_price
                            if total != quantity * unit_price
                            else total
                        )

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

                i += 1

        # Set the display option to show all columns
        pd.set_option("display.max_columns", None)

        # Create a dataframe from the tabulated data
        data_table = pd.DataFrame(tabulated_data, columns=columns)

    return data_table


results = extract_expenses("receipts/Canac")
results.to_excel("receipts/canac_data_scanned.xlsx", index=False)
print("Tabulated data has been saved to 'canac_data_scanned.xlsx'")
print(results)
