from __future__ import annotations

import csv
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Iterable, List, Sequence
from zipfile import ZipFile

NS = {"t": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _column_name_to_index(column: str) -> int:
    """Convert an Excel style column (e.g. 'A', 'AB') to a zero-based index."""
    value = 0
    for char in column:
        if not char.isalpha():
            raise ValueError(f"Invalid column name {column!r}")
        value = value * 26 + (ord(char.upper()) - ord("A") + 1)
    return value - 1


def _load_shared_strings(zf: ZipFile) -> List[str]:
    try:
        with zf.open("xl/sharedStrings.xml") as stream:
            document = ET.parse(stream)
    except KeyError:
        return []

    strings: List[str] = []
    for shared_string in document.getroot().findall("t:si", NS):
        fragments = [node.text or "" for node in shared_string.findall(".//t:t", NS)]
        if not fragments:
            text_node = shared_string.find("t:t", NS)
            fragments = [text_node.text if text_node is not None else ""]
        strings.append("".join(fragments))
    return strings


def _iter_sheet_rows(zf: ZipFile, sheet: str = "xl/worksheets/sheet1.xml") -> Iterable[List[str]]:
    with zf.open(sheet) as stream:
        document = ET.parse(stream)

    shared_strings = _load_shared_strings(zf)
    sheet_data = document.getroot().find("t:sheetData", NS)
    if sheet_data is None:
        return []

    for row in sheet_data.findall("t:row", NS):
        cells: Dict[int, str] = {}
        max_index = -1
        for cell in row.findall("t:c", NS):
            reference = cell.get("r", "")
            column_letters = "".join(filter(str.isalpha, reference))
            if not column_letters:
                continue
            index = _column_name_to_index(column_letters)
            max_index = max(max_index, index)

            value_node = cell.find("t:v", NS)
            raw_value = value_node.text if value_node is not None else ""
            if cell.get("t") == "s" and raw_value:
                raw_value = shared_strings[int(raw_value)]
            cells[index] = raw_value or ""

        if max_index >= 0:
            yield [cells.get(i, "") for i in range(max_index + 1)]


def _clean_numeric(value: str) -> int:
    cleaned = value.strip()
    if not cleaned:
        raise ValueError("Empty numeric value")
    if not cleaned.isdigit():
        raise ValueError(f"Unexpected numeric value: {value!r}")
    return int(cleaned)


def load_house_data(path: Path) -> List[Dict[str, int | str]]:
    with ZipFile(path) as zf:
        rows = list(_iter_sheet_rows(zf))

    if not rows:
        raise ValueError("No rows found in workbook")

    header, *data_rows = rows
    expected_header = ["NO", "NAMA RUMAH", "HARGA", "LB", "LT", "KT", "KM", "GRS"]
    if header != expected_header:
        raise ValueError(f"Unexpected header {header!r}")

    cleaned_rows = []
    for row in data_rows:
        if len(row) < len(expected_header):
            row = row + [""] * (len(expected_header) - len(row))
        try:
            price = _clean_numeric(row[2])
            building_area = _clean_numeric(row[3])
            land_area = _clean_numeric(row[4])
            bedrooms = _clean_numeric(row[5])
            bathrooms = _clean_numeric(row[6])
            garage = _clean_numeric(row[7])
        except ValueError:
            # Skip rows with invalid numeric data
            continue

        cleaned_rows.append(
            {
                "name": row[1].strip(),
                "price": price,
                "building_area": building_area,
                "land_area": land_area,
                "bedrooms": bedrooms,
                "bathrooms": bathrooms,
                "garage": garage,
            }
        )

    unique_rows = []
    seen_keys = set()
    for entry in cleaned_rows:
        key = (
            entry["name"],
            entry["price"],
            entry["building_area"],
            entry["land_area"],
            entry["bedrooms"],
            entry["bathrooms"],
            entry["garage"],
        )
        if key in seen_keys:
            continue
        seen_keys.add(key)
        unique_rows.append(entry)

    return unique_rows


def export_to_csv(rows: Sequence[Dict[str, int | str]], destination: Path) -> None:
    if not rows:
        raise ValueError("No data available to write")

    fieldnames = ["name", "price", "building_area", "land_area", "bedrooms", "bathrooms", "garage"]
    with destination.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)



def main(argv: Sequence[str] | None = None) -> int:
    argv = list(argv or sys.argv[1:])
    if len(argv) not in {0, 2}:
        print("Usage: python clean_data.py [input.xlsx output.csv]", file=sys.stderr)
        return 2

    if argv:
        source = Path(argv[0])
        destination = Path(argv[1])
    else:
        source = Path("DATA RUMAH.xlsx")
        destination = Path("cleaned_house_data.csv")

    rows = load_house_data(source)
    export_to_csv(rows, destination)
    print(f"Wrote {len(rows)} cleaned rows to {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
