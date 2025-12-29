#!/usr/bin/env python3
from __future__ import print_function

import argparse
import difflib
import io
import math
import os
import subprocess
import sys

import xlrd
from xlrd import formatting

BUILTIN_DATE_FORMATS = {
    14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 50, 57, 58,
}


def is_date_cell(book, xf_index):
    if xf_index < 0 or xf_index >= len(book.xf_list):
        return False
    fmt_key = book.xf_list[xf_index].format_key
    if fmt_key in BUILTIN_DATE_FORMATS:
        return True
    fmt = book.format_map.get(fmt_key)
    if fmt is None or not fmt.format_str:
        return False
    return formatting.is_date_format_string(book, fmt.format_str)


def format_float(value, float_format):
    try:
        val = float(value)
    except (TypeError, ValueError):
        return "" if value is None else str(value)
    if float_format:
        return float_format % val
    return repr(val)


def format_date(value, datemode, date_format):
    try:
        val = float(value)
    except (TypeError, ValueError):
        return ""
    if math.isnan(val) or math.isinf(val):
        return ""
    try:
        dt = xlrd.xldate_as_datetime(val, datemode)
    except xlrd.XLDateError:
        return ""
    if date_format:
        return dt.strftime(date_format)
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def format_cell(book, sheet, rowx, colx, date_format, float_format):
    ctype = sheet.cell_type(rowx, colx)
    value = sheet.cell_value(rowx, colx)

    if ctype == xlrd.XL_CELL_TEXT:
        return "" if value is None else str(value)
    if ctype == xlrd.XL_CELL_DATE:
        return format_date(value, book.datemode, date_format)
    if ctype == xlrd.XL_CELL_NUMBER:
        xf_index = sheet.cell_xf_index(rowx, colx)
        if is_date_cell(book, xf_index):
            return format_date(value, book.datemode, date_format)
        return format_float(value, float_format)
    if ctype == xlrd.XL_CELL_BOOLEAN:
        return "TRUE" if value else "FALSE"
    if ctype == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(value, "#ERROR")
    if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return ""
    return "" if value is None else str(value)


def quote_field(value, delimiter):
    if value is None:
        value = ""
    needs_quote = delimiter in value or '"' in value or "\r" in value or "\n" in value
    if not needs_quote:
        return value
    return '"' + value.replace('"', '""') + '"'

def is_empty_csv_row(line, delimiter):
    for ch in line:
        if ch not in (delimiter, '"'):
            return False
    return True


def normalize_csv(text, delimiter):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    while lines and lines[-1] == "":
        lines.pop()
    while lines and is_empty_csv_row(lines[-1], delimiter):
        lines.pop()
    if not lines:
        return ""
    return "\n".join(lines) + "\n"


def sheet_max_cols(sheet):
    max_cols = 0
    for rowx in range(sheet.nrows):
        row_len = sheet.row_len(rowx)
        row_max = 0
        for colx in range(row_len):
            ctype = sheet.cell_type(rowx, colx)
            if ctype not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                row_max = colx + 1
        if row_max > max_cols:
            max_cols = row_max
    if max_cols == 0:
        max_cols = sheet.ncols
    return max_cols


def sheet_to_csv_string(book, sheet, date_format, float_format, delimiter):
    max_cols = sheet_max_cols(sheet)
    buf = io.StringIO()
    for rowx in range(sheet.nrows):
        fields = []
        row_len = sheet.row_len(rowx)
        for colx in range(max_cols):
            if colx >= row_len:
                text = ""
            else:
                text = format_cell(book, sheet, rowx, colx, date_format, float_format)
            fields.append(quote_field(text, delimiter))
        buf.write(delimiter.join(fields) + "\n")
    return buf.getvalue()


def run_xls2csv_output(
    xls2csv, input_path, sheet_id, date_format, float_format, ignore_corruption
):
    cmd = [xls2csv, "-s", str(sheet_id), input_path]
    if ignore_corruption:
        cmd.insert(1, "--ignore-workbook-corruption")
    if date_format:
        cmd.insert(1, date_format)
        cmd.insert(1, "--dateformat")
    if float_format:
        cmd.insert(1, float_format)
        cmd.insert(1, "--floatformat")
    return subprocess.check_output(cmd).decode("utf-8")


def collect_xls_files(paths):
    files = []
    for path in paths:
        if os.path.isdir(path):
            for root, _dirs, filenames in os.walk(path):
                for name in filenames:
                    if name.lower().endswith(".xls"):
                        files.append(os.path.join(root, name))
        elif path.lower().endswith(".xls"):
            files.append(path)
    return sorted(set(files))


def compare_strings(expected, actual, label):
    if expected == actual:
        return True
    diff = difflib.unified_diff(
        expected.splitlines(True),
        actual.splitlines(True),
        fromfile=label + ":python",
        tofile=label + ":xls2csv",
    )
    print("".join(diff))
    return False


def main():
    parser = argparse.ArgumentParser(
        description="Compare xls2csv output with python xlrd output."
    )
    parser.add_argument("--xls2csv", required=True, help="path to xls2csv binary")
    parser.add_argument(
        "--xls-dir",
        action="append",
        default=[],
        help="directory containing xls files (can be repeated)",
    )
    parser.add_argument(
        "--xls",
        action="append",
        default=[],
        help="individual xls file path (can be repeated)",
    )
    parser.add_argument("--dateformat", default="%Y-%m-%d %H:%M:%S")
    parser.add_argument("--floatformat", default="%.15g")
    parser.add_argument(
        "--ignore-workbook-corruption",
        action="store_true",
        help="ignore workbook corruption",
    )
    args = parser.parse_args()

    xls_paths = collect_xls_files(args.xls_dir + args.xls)
    if not xls_paths:
        print("no xls files found")
        return 2

    ok = True
    for path in xls_paths:
        book = xlrd.open_workbook(
            path,
            formatting_info=True,
            ignore_workbook_corruption=args.ignore_workbook_corruption,
        )
        for idx, sheet in enumerate(book.sheets()):
            expected = sheet_to_csv_string(
                book, sheet, args.dateformat, args.floatformat, ","
            )
            actual = run_xls2csv_output(
                args.xls2csv,
                path,
                idx + 1,
                args.dateformat,
                args.floatformat,
                args.ignore_workbook_corruption,
            )
            expected = normalize_csv(expected, ",")
            actual = normalize_csv(actual, ",")
            label = "%s#sheet%d" % (path, idx+1)
            if not compare_strings(expected, actual, label):
                ok = False
                break
        if not ok:
            break

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
