import xml.etree.ElementTree as ET
import openpyxl
import csv
import os


def get_excel_column_values(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        column_values = []

        for row in sheet.iter_rows():
            if len(row) >= 4:
                value = row[3].value
                if value is not None and value != "ID":
                    column_values.append(value.strip())

        return column_values
    except Exception as e:
        raise Exception(f"Error reading Excel: {e}")


def fetch_xml_values(file_path, tags_list):
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        extracted_data = {}

        for elem in root.iter():
            tag = elem.tag.split("}")[-1]
            text = elem.text.strip() if elem.text else ""

            if tag in tags_list:
                if (
                    text
                    and not text.lower().startswith("20")
                    and not text.lower().endswith("date")
                    and text.lower() not in ["", "n/a", "none", "na"]
                ):
                    if tag not in extracted_data:
                        extracted_data[tag] = []
                    extracted_data[tag].append(text)

        return extracted_data
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return {}


def merge_all_data(xml_folder, tags_list):
    merged_data = {tag: [] for tag in tags_list}

    for filename in os.listdir(xml_folder):
        if filename.lower().endswith(".xml"):
            file_path = os.path.join(xml_folder, filename)
            xml_data = fetch_xml_values(file_path, tags_list)

            for tag in tags_list:
                values = xml_data.get(tag, [])
                merged_data[tag].extend(values)

    return merged_data


def write_to_csv(data_dict, output_file_path):
    try:
        headers = list(data_dict.keys())
        max_length = max(len(v) for v in data_dict.values())

        # Prepare rows
        rows = []
        for i in range(max_length):
            row = [data_dict[tag][i] if i < len(data_dict[tag]) else "" for tag in headers]
            rows.append(row)

        # Write to CSV
        with open(output_file_path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            writer.writerows(rows)

        print(f"Successfully created '{output_file_path}'.")
    except Exception as e:
        raise Exception(f"Error writing CSV: {e}")


def main():
    excel_file = "Benchmarking format in Code.xlsx"
    xml_folder = r"C:\Users\AVINANDAN\Desktop\TECH"  # Folder with multiple XML files
    output_file = "output.csv"

    try:
        print(f"Reading tags from '{excel_file}'...")
        tags = get_excel_column_values(excel_file)
        print(f"Found {len(tags)} tags.")

        print(f"Processing XML files from folder '{xml_folder}'...")
        merged_data = merge_all_data(xml_folder, tags)

        print(f"Writing to CSV: {output_file}")
        write_to_csv(merged_data, output_file)

        print("All done.")
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
