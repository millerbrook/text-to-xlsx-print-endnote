import re
import openpyxl

# Define the mapping of column names
COLUMN_MAPPING = {
    'Reference Type': 'Reference Type',
    'Record Number': 'Record Number',
    'Year': 'Year',
    'Title': 'Sender Place',
    'Secondary Author': 'Receiver',
    'Secondary Title': 'Receiver Place',
    'Publisher': 'Publisher',
    'Date': 'date',
    'Type of Work': 'Type of Work',
    'Short Title': 'Collection',
    'Custom 1': 'Custom 1',
    'Custom 2': 'Custom 2',
    'Keywords': 'Keywords',
    'Research Notes': 'Research Notes',
    'Custom 4': 'Digital ID'
}

def parse_txt_to_records(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    records = []
    current_record = {}
    current_array_field = None

    for line in lines:
        line = line.strip()

        # If the line is empty, it indicates the end of a record
        if not line:
            if current_record:
                records.append(current_record)
                current_record = {}
                current_array_field = None
            continue

        # Check if the line contains a colon (indicating a key-value pair)
        if ':' in line:
            key, value = map(str.strip, line.split(':', 1))
            if key in COLUMN_MAPPING:
                current_record[COLUMN_MAPPING[key]] = value
                current_array_field = COLUMN_MAPPING[key]
        else:
            # If the line doesn't contain a colon, treat it as part of an array
            if current_array_field:
                if current_array_field not in current_record:
                    current_record[current_array_field] = []
                if isinstance(current_record[current_array_field], list):
                    current_record[current_array_field].append(line)

    # Add the last record if it exists
    if current_record:
        records.append(current_record)

    # Normalize array fields to comma-separated strings
    for record in records:
        for key, value in record.items():
            if isinstance(value, list):
                record[key] = ', '.join(value)

    return records

def write_records_to_xlsx(records, output_file):
    # Create a new workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write the header row
    headers = list(COLUMN_MAPPING.values())
    sheet.append(headers)

    # Write the records
    for record in records:
        row = [record.get(header, None) for header in headers]
        sheet.append(row)

    # Save the workbook
    workbook.save(output_file)

def main():
    input_file = 'Amsterdam.txt'
    output_file = 'converted_records.xlsx'

    # Parse the text file into records
    records = parse_txt_to_records(input_file)

    # Write the records to an Excel file
    write_records_to_xlsx(records, output_file)

    print(f"Conversion complete. The records have been saved to '{output_file}'.")

if __name__ == '__main__':
    main()