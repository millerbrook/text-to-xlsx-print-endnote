import re
import pandas as pd

# Define valid column headings
VALID_COLUMNS = [
    'Reference Type', 'Record Number', 'Author', 'Year', 'Title',
    'Secondary Author', 'Secondary Title', 'Publisher', 'Date',
    'Type of Work', 'Short Title', 'Custom 1', 'Custom 2', 'Custom 3', 'Custom 4',
    'Keywords', 'Abstract', 'Notes', 'Research Notes', "'File' Attachments"
]

# Define column renaming rules
COLUMN_RENAMING = {
    'Author': 'Sender',
    'Title': 'Sender Place',
    'Secondary Author': 'Receiver',
    'Secondary Title': 'Receiver Place',
    'Date': 'date',
    'Short Title': 'collection',
    'Custom 4': 'Digital ID'
}

def parse_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    records = []
    current_record = {}
    current_key = None
    empty_line_count = 0  # Track consecutive empty lines

    for line in lines:
        line = line.strip()

        # Check for empty lines
        if not line:
            empty_line_count += 1
            if empty_line_count == 2:  # End of a record
                if current_record:
                    records.append(current_record)
                    current_record = {}
                    current_key = None
                empty_line_count = 0
            continue
        else:
            empty_line_count = 0  # Reset empty line count if a non-empty line is encountered

        # Match lines with a valid column header
        match = re.match(r'^([\w\s\'"]+):\s*(.*)', line)
        if match:
            key = match.group(1).strip()
            value = match.group(2).strip()

            # Map 'Notes' to 'Research Notes'
            if key == 'Notes':
                key = 'Research Notes'

            # Only process valid column headings
            if key in VALID_COLUMNS:
                current_key = key
                if key in current_record:
                    # If the key already exists, append the new value
                    if isinstance(current_record[key], list):
                        current_record[key].append(value)
                    else:
                        current_record[key] = [current_record[key], value]
                else:
                    current_record[key] = value
        else:
            # If no valid column header is found, treat it as a continuation of the current key
            if current_key and current_key in current_record:
                if isinstance(current_record[current_key], list):
                    current_record[current_key].append(line)
                else:
                    current_record[current_key] = [current_record[current_key], line]

    # Add the last record if the file doesn't end with two blank lines
    if current_record:
        records.append(current_record)

    # Normalize multi-line fields to comma-separated strings
    for record in records:
        for key, value in record.items():
            if isinstance(value, list):
                record[key] = ', '.join(value)

    return records

def main():
    # Path to the input file
    input_file = 'Berlin StaBi SAF.txt'

    # Parse the file
    records = parse_file(input_file)

    # Convert to a DataFrame
    df = pd.DataFrame(records)

    # Rename columns based on the COLUMN_RENAMING dictionary
    df.rename(columns=COLUMN_RENAMING, inplace=True)

    # Drop the "Reference Type" column if it exists
    if 'Reference Type' in df.columns:
        df.drop(columns=['Reference Type'], inplace=True)

    # Ensure single values are not stored as arrays
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x if not isinstance(x, list) else (x[0] if len(x) == 1 else x))

    # Save to an Excel file
    output_file = 'converted_records.xlsx'
    df.to_excel(output_file, index=False)
    print(f"Data has been parsed and saved to {output_file}")

if __name__ == '__main__':
    main()