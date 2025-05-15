import argparse
import logging
import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import csv
import pandas as pd  # Using pandas for more robust CSV handling

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def setup_argparse():
    """
    Sets up the argument parser for the command-line interface.
    """
    parser = argparse.ArgumentParser(description="Sanitizes spreadsheet files by removing macros and hidden sheets.")
    parser.add_argument("input_file", help="The input spreadsheet file (xlsx, ods, or csv).")
    parser.add_argument("output_file", help="The output sanitized spreadsheet file.")
    parser.add_argument("--remove-macros", action="store_true", help="Remove VBA macros (xlsx only).")
    parser.add_argument("--remove-hidden-sheets", action="store_true", help="Remove hidden sheets (xlsx and ods).")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite the output file if it exists.")

    return parser.parse_args()


def sanitize_xlsx(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
    """
    Sanitizes an XLSX file.

    Args:
        input_file (str): Path to the input XLSX file.
        output_file (str): Path to the output sanitized XLSX file.
        remove_macros (bool): Whether to remove VBA macros.
        remove_hidden_sheets (bool): Whether to remove hidden sheets.
        overwrite (bool): Whether to overwrite the output file if it exists.
    """
    try:
        if os.path.exists(output_file) and not overwrite:
            logging.error(f"Output file '{output_file}' already exists. Use --overwrite to replace it.")
            return False

        with zipfile.ZipFile(input_file, 'r') as zin:
            with zipfile.ZipFile(output_file, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    filename = item.filename

                    # Remove VBA Project (macros)
                    if remove_macros and filename == 'xl/vbaProject.bin':
                        logging.info("Removing VBA macros.")
                        continue

                    # Remove hidden sheets from workbook.xml
                    if remove_hidden_sheets and filename == 'xl/workbook.xml':
                        logging.info("Removing hidden sheets.")
                        content = zin.read(filename)
                        try:
                            tree = ET.fromstring(content)
                            # Define XML namespaces
                            namespaces = {'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

                            # Find all sheets
                            sheets = tree.findall('.//xmlns:sheet', namespaces)
                            sheets_to_remove = []

                            for sheet in sheets:
                                if sheet.get('state') == 'hidden' or sheet.get('state') == 'veryHidden':
                                    sheets_to_remove.append(sheet.get('sheetId'))
                                    logging.info(f"Removing sheet with id {sheet.get('sheetId')}")

                            # remove the sheet
                            for sheet in sheets:
                                if sheet.get('sheetId') in sheets_to_remove:
                                    sheet.getparent().remove(sheet)
                            content = ET.tostring(tree, encoding='utf8').decode('utf8')
                        except Exception as e:
                            logging.error(f"Error parsing workbook.xml: {e}")
                            zout.writestr(filename, zin.read(filename)) # Write original file if parsing failed
                            continue

                        zout.writestr(filename, content)
                        continue #skip writing the original workbook.xml file


                    # Copy everything else
                    buffer = zin.read(filename)
                    zout.writestr(item, buffer)
        logging.info(f"Successfully sanitized XLSX file. Output: {output_file}")
        return True

    except FileNotFoundError:
        logging.error(f"Input file '{input_file}' not found.")
        return False
    except zipfile.BadZipFile:
        logging.error(f"Input file '{input_file}' is not a valid XLSX file.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        return False


def sanitize_ods(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
    """
    Sanitizes an ODS file. ODS files are similar to XLSX but handle hidden sheets differently.
    Macros in ODS are more complex and not directly removable with this simple approach.

    Args:
        input_file (str): Path to the input ODS file.
        output_file (str): Path to the output sanitized ODS file.
        remove_macros (bool): Not directly supported for ODS (indicates intention but doesn't remove macros).
        remove_hidden_sheets (bool): Whether to remove hidden sheets.
        overwrite (bool): Whether to overwrite the output file if it exists.
    """
    try:
        if os.path.exists(output_file) and not overwrite:
            logging.error(f"Output file '{output_file}' already exists. Use --overwrite to replace it.")
            return False

        if remove_macros:
            logging.warning("Macro removal for ODS files is not fully supported with this simple approach.")

        with zipfile.ZipFile(input_file, 'r') as zin:
            with zipfile.ZipFile(output_file, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    filename = item.filename

                    # Remove hidden sheets from content.xml
                    if remove_hidden_sheets and filename == 'content.xml':
                        logging.info("Removing hidden sheets from content.xml.")
                        content = zin.read(filename)
                        try:
                            tree = ET.fromstring(content)
                            namespaces = {'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                                        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'}

                            tables = tree.findall('.//table:table', namespaces)
                            tables_to_remove = []
                            for table in tables:
                                if table.get('table:display', default='true', namespaces=namespaces) == 'false':  # Check for hidden tables
                                    table_name = table.get('table:name', namespaces=namespaces)
                                    tables_to_remove.append(table)
                                    logging.info(f"Removing table: {table_name}")

                            for table in tables_to_remove:
                                table.getparent().remove(table)

                            content = ET.tostring(tree, encoding='utf8').decode('utf8')

                        except Exception as e:
                            logging.error(f"Error parsing content.xml: {e}")
                            zout.writestr(filename, zin.read(filename)) # Write original file if parsing failed
                            continue

                        zout.writestr(filename, content)
                        continue # skip writing the original content.xml file

                    # Copy everything else
                    buffer = zin.read(filename)
                    zout.writestr(item, buffer)

        logging.info(f"Successfully sanitized ODS file. Output: {output_file}")
        return True

    except FileNotFoundError:
        logging.error(f"Input file '{input_file}' not found.")
        return False
    except zipfile.BadZipFile:
        logging.error(f"Input file '{input_file}' is not a valid ODS file.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        return False


def sanitize_csv(input_file, output_file, overwrite):
    """
    Sanitizes a CSV file. CSV sanitization primarily focuses on ensuring proper formatting and encoding.
    This implementation provides a basic example. More robust sanitization might involve data type validation,
    escaping special characters, and handling different delimiters and quoting styles.

    Args:
        input_file (str): Path to the input CSV file.
        output_file (str): Path to the output sanitized CSV file.
        overwrite (bool): Whether to overwrite the output file if it exists.
    """
    try:
        if os.path.exists(output_file) and not overwrite:
            logging.error(f"Output file '{output_file}' already exists. Use --overwrite to replace it.")
            return False

        # Use pandas for more robust CSV reading and writing
        df = pd.read_csv(input_file)  # Let pandas handle encoding detection and quoting

        # Simple example: remove rows with any missing values
        df = df.dropna()  # Remove rows with missing values. Good for sanitization
        df.to_csv(output_file, index=False) # Write the DataFrame to a CSV file

        logging.info(f"Successfully sanitized CSV file. Output: {output_file}")
        return True

    except FileNotFoundError:
        logging.error(f"Input file '{input_file}' not found.")
        return False
    except pd.errors.EmptyDataError:
         logging.error(f"Input file '{input_file}' is empty.")
         return False
    except pd.errors.ParserError:
        logging.error(f"Input file '{input_file}' is not a valid CSV file.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        return False


def main():
    """
    Main function to parse arguments and call the appropriate sanitization function.
    """
    args = setup_argparse()
    input_file = args.input_file
    output_file = args.output_file
    remove_macros = args.remove_macros
    remove_hidden_sheets = args.remove_hidden_sheets
    overwrite = args.overwrite

    # Determine file type based on extension
    file_extension = os.path.splitext(input_file)[1].lower()

    if file_extension == '.xlsx':
        if not sanitize_xlsx(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
            sys.exit(1) # Exit with error code if sanitization failed
    elif file_extension == '.ods':
        if not sanitize_ods(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
            sys.exit(1) # Exit with error code if sanitization failed
    elif file_extension == '.csv':
        if not sanitize_csv(input_file, output_file, overwrite):
            sys.exit(1)  # Exit with error code if sanitization failed
    else:
        logging.error("Unsupported file type.  Supported types are .xlsx, .ods, and .csv")
        sys.exit(1) # Exit with error code for unsupported file type

    sys.exit(0)  # Exit with success code


if __name__ == "__main__":
    main()