import argparse
import logging
import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import csv
import io

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
OFFICE_NS = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
TABLE_NS = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
TEXT_NS = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'


def setup_argparse():
    parser = argparse.ArgumentParser(description="Sanitizes spreadsheet files by removing macros and hidden sheets.")
    parser.add_argument("input_file", help="The input spreadsheet file (xlsx, ods, or csv).")
    parser.add_argument("output_file", help="The output sanitized spreadsheet file. It must not resolve to the input file; --overwrite does not permit in-place work.")
    parser.add_argument("--remove-macros", action="store_true", help="Remove VBA macros (xlsx only).")
    parser.add_argument("--remove-hidden-sheets", action="store_true", help="Remove hidden sheets (xlsx and ods).")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite the output file if it exists. This consents to replacing an existing output file, not to in-place work on the input file.")
    return parser.parse_args()


def _same_path(a, b):
    return os.path.realpath(a) == os.path.realpath(b)


def _read_zip_entries(input_file):
    with zipfile.ZipFile(input_file, 'r') as zin:
        return [(item, zin.read(item.filename)) for item in zin.infolist()]


def _write_zip_entries(output_file, entries):
    with zipfile.ZipFile(output_file, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item, data in entries:
            zout.writestr(item, data)


def _remove_hidden_sheets_xlsx(entries):
    """Remove hidden sheets. Returns (entries, removed_count). Raises on failure."""
    by_name = {item.filename: data for item, data in entries}
    if 'xl/workbook.xml' not in by_name:
        raise ValueError("xl/workbook.xml not found")
    ET.register_namespace('', MAIN_NS)
    ET.register_namespace('r', REL_NS)
    tree = ET.fromstring(by_name['xl/workbook.xml'])
    sheets_el = tree.find(f'{{{MAIN_NS}}}sheets')
    if sheets_el is None:
        raise ValueError("sheets element not found in workbook.xml")
    hidden = []
    for sheet in list(sheets_el):
        if sheet.get('state') in ('hidden', 'veryHidden'):
            hidden.append(sheet)
    if not hidden:
        return entries, 0
    rels_name = 'xl/_rels/workbook.xml.rels'
    if rels_name not in by_name:
        raise ValueError("workbook relationships not found")
    ET.register_namespace('', PKG_REL_NS)
    rels_tree = ET.fromstring(by_name[rels_name])
    rel_by_id = {}
    for rel in rels_tree:
        rel_by_id[rel.get('Id')] = rel
    ct_name = '[Content_Types].xml'
    ct_tree = None
    if ct_name in by_name:
        ET.register_namespace('', CT_NS)
        ct_tree = ET.fromstring(by_name[ct_name])
    removed_parts = []
    for sheet in hidden:
        rid = sheet.get(f'{{{REL_NS}}}id')
        rel = rel_by_id.get(rid)
        if rel is None:
            raise ValueError(f"hidden sheet relationship {rid} not found")
        target = rel.get('Target')
        if target.startswith('/'):
            part = target.lstrip('/')
        else:
            part = 'xl/' + target
        part = os.path.normpath(part).replace(os.sep, '/')
        if part not in by_name:
            raise ValueError(f"hidden sheet part {part} not found")
        removed_parts.append(part)
        sheets_el.remove(sheet)
        rels_tree.remove(rel)
        logging.info(f"Removing hidden sheet: {sheet.get('name')}")
    new_entries = []
    for item, data in entries:
        name = item.filename
        if name in removed_parts:
            continue
        if name == 'xl/workbook.xml':
            data = ET.tostring(tree, encoding='utf8')
        elif name == rels_name:
            data = ET.tostring(rels_tree, encoding='utf8')
        elif name == ct_name and ct_tree is not None:
            for override in list(ct_tree):
                if override.get('PartName') in ('/' + p for p in removed_parts):
                    ct_tree.remove(override)
            data = ET.tostring(ct_tree, encoding='utf8')
        new_entries.append((item, data))
    return new_entries, len(hidden)


def _remove_macros_xlsx(entries):
    """Remove the VBA project and its references. Returns (entries, removed_count)."""
    names = [item.filename for item, data in entries]
    if 'xl/vbaProject.bin' not in names:
        return entries, 0
    new_entries = []
    for item, data in entries:
        name = item.filename
        if name == 'xl/vbaProject.bin':
            logging.info("Removing VBA macros.")
            continue
        if name == '[Content_Types].xml':
            ET.register_namespace('', CT_NS)
            ct_tree = ET.fromstring(data)
            for override in list(ct_tree):
                if override.get('PartName') == '/xl/vbaProject.bin':
                    ct_tree.remove(override)
                ct = override.get('ContentType') or ''
                if 'macroEnabled' in ct:
                    override.set('ContentType', ct.replace(
                        'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'))
            data = ET.tostring(ct_tree, encoding='utf8')
        elif name == 'xl/_rels/workbook.xml.rels':
            ET.register_namespace('', PKG_REL_NS)
            rels_tree = ET.fromstring(data)
            for rel in list(rels_tree):
                if 'vbaProject' in (rel.get('Type') or '') or 'vbaProject' in (rel.get('Target') or ''):
                    rels_tree.remove(rel)
            data = ET.tostring(rels_tree, encoding='utf8')
        new_entries.append((item, data))
    return new_entries, 1


def sanitize_xlsx(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
    try:
        if _same_path(input_file, output_file):
            logging.error("Output file resolves to the input file; refusing in-place work, even with --overwrite.")
            return False
        if os.path.exists(output_file) and not overwrite:
            logging.error(f"Output file '{output_file}' already exists. Use --overwrite to replace it.")
            return False
        entries = _read_zip_entries(input_file)
        removed = 0
        if remove_hidden_sheets:
            try:
                entries, n = _remove_hidden_sheets_xlsx(entries)
            except Exception as e:
                logging.error(f"Could not remove hidden sheets: {e}")
                return False
            removed += n
        if remove_macros:
            entries, n = _remove_macros_xlsx(entries)
            removed += n
        if removed == 0:
            logging.error("Nothing was removed; the file was not sanitized.")
            return False
        _write_zip_entries(output_file, entries)
        logging.info(f"Successfully sanitized XLSX file. Output: {output_file} ({removed} item(s) removed)")
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


def _remove_hidden_tables_ods(entries):
    by_name = {item.filename: data for item, data in entries}
    if 'content.xml' not in by_name:
        raise ValueError("content.xml not found")
    ET.register_namespace('office', OFFICE_NS)
    ET.register_namespace('table', TABLE_NS)
    ET.register_namespace('text', TEXT_NS)
    tree = ET.fromstring(by_name['content.xml'])
    parent = {child: p for p in tree.iter() for child in p}
    hidden = []
    for table in tree.iter(f'{{{TABLE_NS}}}table'):
        if table.get(f'{{{TABLE_NS}}}display') == 'false':
            hidden.append(table)
    if not hidden:
        return entries, 0
    for table in hidden:
        p = parent.get(table)
        if p is None:
            raise ValueError("hidden table has no parent element")
        logging.info(f"Removing hidden table: {table.get(f'{{{TABLE_NS}}}name')}")
        p.remove(table)
    new_entries = []
    for item, data in entries:
        if item.filename == 'content.xml':
            data = ET.tostring(tree, encoding='utf8')
        new_entries.append((item, data))
    return new_entries, len(hidden)


def sanitize_ods(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
    try:
        if _same_path(input_file, output_file):
            logging.error("Output file resolves to the input file; refusing in-place work, even with --overwrite.")
            return False
        if os.path.exists(output_file) and not overwrite:
            logging.error(f"Output file '{output_file}' already exists. Use --overwrite to replace it.")
            return False
        if remove_macros:
            logging.warning("Macro removal for ODS files is not supported by this tool.")
        entries = _read_zip_entries(input_file)
        removed = 0
        if remove_hidden_sheets:
            try:
                entries, n = _remove_hidden_tables_ods(entries)
            except Exception as e:
                logging.error(f"Could not remove hidden sheets: {e}")
                return False
            removed += n
        if removed == 0:
            logging.error("Nothing was removed; the file was not sanitized.")
            return False
        _write_zip_entries(output_file, entries)
        logging.info(f"Successfully sanitized ODS file. Output: {output_file} ({removed} item(s) removed)")
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
    try:
        if _same_path(input_file, output_file):
            logging.error("Output file resolves to the input file; refusing in-place work, even with --overwrite.")
            return False
        if os.path.exists(output_file) and not overwrite:
            logging.error(f"Output file '{output_file}' already exists. Use --overwrite to replace it.")
            return False
        with open(input_file, 'r', encoding='utf-8', newline='') as f:
            raw = f.read()
        if not raw.strip():
            logging.error(f"Input file '{input_file}' is empty.")
            return False
        rows = list(csv.reader(io.StringIO(raw)))
        if not rows:
            logging.error(f"Input file '{input_file}' is empty.")
            return False
        neutralized = 0
        out_rows = []
        for row in rows:
            new_row = []
            for cell in row:
                if cell[:1] in ('=', '+', '@'):
                    new_row.append("'" + cell)
                    neutralized += 1
                else:
                    new_row.append(cell)
            out_rows.append(new_row)
        if neutralized == 0:
            logging.error("No formula cells were neutralized; nothing was changed.")
            return False
        with open(output_file, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(out_rows)
        logging.info(f"Successfully sanitized CSV file. Output: {output_file} ({len(out_rows)} row(s) kept, {neutralized} formula cell(s) neutralized)")
        return True
    except FileNotFoundError:
        logging.error(f"Input file '{input_file}' not found.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        return False


def main():
    args = setup_argparse()
    input_file = args.input_file
    output_file = args.output_file
    remove_macros = args.remove_macros
    remove_hidden_sheets = args.remove_hidden_sheets
    overwrite = args.overwrite
    if _same_path(input_file, output_file):
        logging.error("Output file resolves to the input file; refusing in-place work, even with --overwrite.")
        sys.exit(1)
    file_extension = os.path.splitext(input_file)[1].lower()
    if file_extension == '.xlsx':
        if not sanitize_xlsx(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
            sys.exit(1)
    elif file_extension == '.ods':
        if not sanitize_ods(input_file, output_file, remove_macros, remove_hidden_sheets, overwrite):
            sys.exit(1)
    elif file_extension == '.csv':
        if not sanitize_csv(input_file, output_file, overwrite):
            sys.exit(1)
    else:
        logging.error("Unsupported file type.  Supported types are .xlsx, .ods, and .csv")
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    main()
