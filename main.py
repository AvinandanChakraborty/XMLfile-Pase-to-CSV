import xml.etree.ElementTree as ET
import openpyxl
import csv


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
    except FileNotFoundError:
        raise FileNotFoundError(f"Error: The file '{file_path}' was not found.")
    except Exception as e:
        raise Exception(f"Error while reading Excel: {e}")


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

        # Only keep tags that have at least one valid value
        extracted_data = {k: v for k, v in extracted_data.items() if v}

        return extracted_data
    except FileNotFoundError:
        raise FileNotFoundError(f"Error: The file '{file_path}' was not found.")
    except ET.ParseError as e:
        raise Exception(f"XML parsing error: {e}")
    except Exception as e:
        raise Exception(f"Error while processing XML: {e}")


def write_to_csv(data_dict, output_file_path):
    try:
        if not data_dict:
            print("No valid data found to write.")
            return

        headers = list(data_dict.keys())
        values = [data_dict[tag][0] if data_dict[tag] else "" for tag in headers]

        with open(output_file_path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)  # First row: tags
            writer.writerow(values)   # Second row: first value of each tag

        print(f"Successfully created '{output_file_path}'.")
    except Exception as e:
        raise Exception(f"Error while writing CSV: {e}")


def main():
    excel_file = "Benchmarking format in Code.xlsx"
    xml_file = "BRSR_926377_02092023090138_WEB.xml"
    output_file = "output.csv"

    try:
        print(f"Reading tags from '{excel_file}'...")
        excel_tags = get_excel_column_values(excel_file)
        print(f"Found {len(excel_tags)} tags from Excel.")

        print(f"Fetching XML data from '{xml_file}'...")
        xml_data = fetch_xml_values(xml_file, excel_tags)
        print(f"Found {len(xml_data)} valid tags with values.")

        print(f"Writing cleaned data to '{output_file}'...")
        write_to_csv(xml_data, output_file)

        print("Script finished successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
