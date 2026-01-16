import datetime as dt
import random
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET


random.seed(42)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"main": NS_MAIN, "rel": NS_REL, "pkg": NS_PKG}

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def col_index_to_letter(idx: int) -> str:
    letters = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def col_letter_to_index(letter: str) -> int:
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - 64)
    return idx


def parse_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: List[str] = []
    for si in root.findall("main:si", NS):
        parts = []
        for t in si.findall(".//main:t", NS):
            parts.append(t.text or "")
        strings.append("".join(parts))
    return strings


def decode_cell(cell: ET.Element, shared_strings: List[str]) -> str:
    t = cell.attrib.get("t")
    v = cell.find("main:v", NS)
    if t == "s" and v is not None:
        idx = int(v.text)
        return shared_strings[idx] if idx < len(shared_strings) else ""
    if t == "inlineStr":
        t_node = cell.find(".//main:t", NS)
        return t_node.text or "" if t_node is not None else ""
    if v is not None:
        return v.text or ""
    return ""


def parse_sheet_meta(
    zf: zipfile.ZipFile, sheet_path: str, shared_strings: List[str]
) -> Tuple[List[str], int, str]:
    headers: List[str] = []
    max_row = 1
    max_col_letter = None
    dimension_ref = None
    with zf.open(sheet_path) as fh:
        context = ET.iterparse(fh, events=("start", "end"))
        for event, elem in context:
            tag = elem.tag.split("}")[-1]
            if event == "start" and tag == "dimension":
                dimension_ref = elem.attrib.get("ref")
            if event == "end" and tag == "row":
                r_idx = int(elem.attrib.get("r", len(headers) + 1))
                if r_idx == 1:
                    row_cells = elem.findall("main:c", NS)
                    sorted_cells = sorted(
                        row_cells,
                        key=lambda c: col_letter_to_index(
                            re.match(r"([A-Z]+)", c.attrib.get("r", "")).group(1)
                            if re.match(r"([A-Z]+)", c.attrib.get("r", ""))
                            else "A"
                        ),
                    )
                    headers = [decode_cell(c, shared_strings) for c in sorted_cells]
                max_row = max(max_row, r_idx)
                if r_idx >= 2:
                    break
    if dimension_ref:
        end_ref = dimension_ref.split(":")[1] if ":" in dimension_ref else dimension_ref
        m = re.match(r"([A-Z]+)([0-9]+)", end_ref)
        if m:
            max_col_letter, max_row = m.group(1), int(m.group(2))
    if not max_col_letter and headers:
        max_col_letter = col_index_to_letter(len(headers))
    return headers, max_row, max_col_letter or "A"


def build_row_xml(
    row_idx: int, values: List, num_flags: List[bool], col_letters: List[str]
) -> str:
    parts = [f'<row r="{row_idx}">']
    for col_letter, val, is_num in zip(col_letters, values, num_flags):
        ref = f"{col_letter}{row_idx}"
        if is_num:
            parts.append(f'<c r="{ref}"><v>{val}</v></c>')
        else:
            text = escape(str(val) if val is not None else "")
            parts.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
    parts.append("</row>\n")
    return "".join(parts)


# Synthetic data pools
COMPANIES = [
    "Aurora Foods",
    "Pacific Harvest",
    "Golden Fields Co",
    "Summit Brands",
    "TerraNova Foods",
    "Blue Mesa Group",
    "Northern Lights Foods",
    "Silver Crest Foods",
    "Frontier Organics",
    "Oak & Pine Foods",
]
FRANCHISES = [
    "Everyday Favorites",
    "Chef Select",
    "Market Classics",
    "Fresh Corner",
    "Heritage Kitchen",
]
ATTRIBUTES_POOL = [
    "Category Int Fresh",
    "Premium Line",
    "Family Pack",
    "Organic Range",
    "Value Line",
]
GEOS = [
    "Total US - Multi Outlet+",
    "Total US - Grocery",
    "Northeast",
    "Southeast",
    "Midwest",
    "West",
]
PACKAGE_TYPES = ["Can", "Pouch", "Bottle", "Box", "Jar", "Tray"]
CATEGORIES = {
    "Snacks": ["Chips", "Crackers", "Nuts"],
    "Beverages": ["Sparkling Water", "Juice", "Energy Drinks"],
    "Dairy": ["Yogurt", "Cheese", "Butter"],
    "Bakery": ["Bread", "Cookies", "Cakes"],
    "Frozen": ["Frozen Meals", "Ice Cream", "Frozen Vegetables"],
    "Pantry": ["Pasta", "Sauces", "Rice"],
    "Breakfast": ["Cereal", "Oatmeal", "Granola"],
    "Protein": ["Chicken", "Beef", "Plant-Based"],
    "Produce": ["Salad Kits", "Fresh Berries", "Cut Fruit"],
}
SIZES = ["8 oz", "12 oz", "16 oz", "24 oz", "32 oz", "48 oz", "64 oz"]
MONTH_NAMES = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]


def make_week_rows(count: int) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    base_date = dt.date(2025, 1, 5)
    for i in range(count):
        idx = i + 1
        date_value = base_date + dt.timedelta(days=7 * i)
        rows.append(
            {"Time": f"Week Ending {date_value.strftime('%m-%d-%y')}", "Week": idx}
        )
    return rows


def make_brand_rows(count: int) -> List[Dict[str, str]]:
    names = [
        "Evergreen",
        "Solstice",
        "Cascade",
        "Lumen",
        "Harbor",
        "Mesa",
        "Redwood",
        "Juniper",
        "Arroyo",
        "Summit",
    ]
    suffixes = ["Foods", "Kitchen", "Harvest", "Market"]
    rows: List[Dict[str, str]] = []
    for _ in range(count):
        brand_name = f"{random.choice(names)} {random.choice(suffixes)}"
        rows.append({"Brand": brand_name.upper(), "Name": brand_name})
    return rows


def make_category_rows(count: int, brands: List[Dict[str, str]]) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    barcode_base = 7600000000000
    for i in range(count):
        category = random.choice(list(CATEGORIES.keys()))
        subcat = random.choice(CATEGORIES[category])
        brand = random.choice(brands)["Brand"]
        size = random.choice(SIZES)
        barcode = str(barcode_base + i)
        product = f"{brand} {subcat} {size} - {barcode}"
        rows.append({"Product": product, "Category": category, "Subcategory": subcat})
    return rows


def rand_money(low=100, high=5000) -> float:
    return round(random.uniform(low, high), 2)


def rand_units(low=1, high=1200) -> float:
    return round(random.uniform(low, high), 2)


def rand_ratio() -> float:
    return round(random.uniform(0.05, 0.95), 4)


def rand_price() -> float:
    return round(random.uniform(0.5, 25), 3)


def build_source_row(
    brand_rows: List[Dict[str, str]],
    category_rows: List[Dict[str, str]],
    week_rows: List[Dict[str, object]],
) -> List[object]:
    brand_entry = random.choice(brand_rows)
    cat_entry = random.choice(category_rows)
    week_entry = random.choice(week_rows)
    month_num = random.randint(1, 12)
    month_name = MONTH_NAMES[month_num - 1]
    year_val = random.choice([2023, 2024, 2025])
    total_ounces = round(random.uniform(4, 96), 1)
    price_unit = rand_price()
    units = rand_units(5, 800)
    dollars = round(units * price_unit, 2)
    price_unit_prev = max(0.1, price_unit * random.uniform(0.85, 1.15))
    units_prev = rand_units(5, 800)
    dollars_prev = round(units_prev * price_unit_prev, 2)
    stores_total = random.randint(25000, 150000)
    stores_selling = random.randint(max(500, stores_total // 8), stores_total)
    items_per_store = round(random.uniform(0.5, 8.0), 3)
    avg_units_per_store = round(units / max(1, stores_selling) * random.uniform(8, 18), 4)
    avg_dollars_per_store = round(
        dollars / max(1, stores_selling) * random.uniform(8, 18), 4
    )
    return [
        random.choice(COMPANIES),  # Parent Company-Int Fresh
        random.choice(FRANCHISES),  # Brand Franchise-Int Fresh
        random.choice(ATTRIBUTES_POOL),  # Integrated Fresh Attributes
        "MULO PLUS",  # MULOPLUS
        cat_entry["Product"],  # Product
        random.choice(GEOS),  # Geography
        week_entry["Time"],  # Time
        week_entry["Week"],  # Week
        month_num,  # Mes#
        month_name,  # Mes name
        f"{month_num}. {month_name[:3]}",  # Mes code
        year_val,  # Year
        brand_entry["Brand"],  # Brand-Int Fresh Value
        total_ounces,  # Total Ounces
        random.choice(PACKAGE_TYPES),  # Package Type-Int Fresh Value
        cat_entry["Category"],  # Category-Int Fresh Value
        cat_entry["Subcategory"],  # Subcategory-Int Fresh Value
        dollars,  # Dollar Sales
        dollars_prev,  # Dollar Sales Year Ago
        units,  # Unit Sales
        units_prev,  # Unit Sales Year Ago
        rand_ratio(),  # ACV Weighted Distribution
        rand_ratio(),  # ACV Weighted Distribution Year Ago
        items_per_store,  # Avg Weekly Items per Store Selling
        stores_total,  # Number of Stores
        round(stores_selling * random.uniform(0.5, 1.2), 3),  # Number of Stores Selling
        price_unit,  # Price per Unit
        round(price_unit_prev, 3),  # Price per Unit Year Ago
        avg_units_per_store,  # Avg Weekly Units per Store Selling
        avg_dollars_per_store,  # Avg Weekly Dollars per Store Selling
        brand_entry["Brand"],  # Brand SM
        cat_entry["Category"],  # Category SM
        cat_entry["Subcategory"],  # Subcategory SM
    ]


def write_table_sheet(
    out_path: Path, headers: List[str], rows: Iterable[List[object]], num_flags: List[bool]
) -> None:
    col_letters = [col_index_to_letter(i + 1) for i in range(len(headers))]
    rows = list(rows)
    total_rows = 1 + len(rows)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(
            '<worksheet xmlns="{0}" xmlns:r="{1}">'.format(NS_MAIN, NS_REL)
        )
        f.write(
            '<dimension ref="A1:{0}{1}"/>'.format(col_letters[-1], total_rows)
        )
        f.write("<sheetData>")
        f.write(build_row_xml(1, headers, [False] * len(headers), col_letters))
        for idx, row in enumerate(rows, start=2):
            f.write(build_row_xml(idx, row, num_flags, col_letters))
        f.write("</sheetData></worksheet>")


def write_source_sheet(
    out_path: Path,
    headers: List[str],
    num_flags: List[bool],
    row_count: int,
    brand_rows: List[Dict[str, str]],
    category_rows: List[Dict[str, str]],
    week_rows: List[Dict[str, object]],
) -> None:
    col_letters = [col_index_to_letter(i + 1) for i in range(len(headers))]
    with out_path.open("w", encoding="utf-8", newline="") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(
            '<worksheet xmlns="{0}" xmlns:r="{1}">'.format(NS_MAIN, NS_REL)
        )
        f.write(
            '<dimension ref="A1:{0}{1}"/>'.format(col_letters[-1], row_count)
        )
        f.write("<sheetData>")
        f.write(build_row_xml(1, headers, [False] * len(headers), col_letters))
        data_rows = row_count - 1
        for idx in range(2, row_count + 1):
            row = build_source_row(brand_rows, category_rows, week_rows)
            f.write(build_row_xml(idx, row, num_flags, col_letters))
        f.write("</sheetData></worksheet>")


def main():
    src_path = Path("Salida") / "00. SM-SourceOfTruth.xlsx"
    dst_path = src_path.with_name(src_path.stem + "_fake.xlsx")

    with zipfile.ZipFile(src_path, "r") as zf:
        shared_strings = parse_shared_strings(zf)
        rels_tree = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_tree.findall("pkg:Relationship", NS)}
        wb_tree = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet_info: Dict[str, Dict[str, object]] = {}
        for sheet in wb_tree.findall("main:sheets/main:sheet", NS):
            name = sheet.attrib["name"]
            rId = sheet.attrib["{" + NS_REL + "}id"]
            target = rel_map[rId].lstrip("/")
            sheet_path = ("xl/" + target) if not target.startswith("xl/") else target
            headers, max_row, max_col_letter = parse_sheet_meta(zf, sheet_path, shared_strings)
            sheet_info[name] = {
                "path": sheet_path,
                "headers": headers,
                "max_row": max_row,
                "max_col_letter": max_col_letter,
            }

    # Build fake dictionaries sized like the originals
    week_data = make_week_rows(max(sheet_info["Week Dictionary"]["max_row"] - 1, 0))
    brand_data = make_brand_rows(max(sheet_info["Brand Dictionary"]["max_row"] - 1, 0))
    category_data = make_category_rows(
        max(sheet_info["Category Dictionary"]["max_row"] - 1, 0), brand_data
    )

    source_headers: List[str] = sheet_info["Source of Truth"]["headers"]  # type: ignore
    source_row_count: int = sheet_info["Source of Truth"]["max_row"]  # type: ignore
    source_num_flags = [
        False,
        False,
        False,
        False,
        False,  # 1-5
        False,
        False,
        True,
        True,
        False,  # 6-10
        False,
        True,
        False,
        True,
        False,  # 11-15
        False,
        False,
        False,
        True,
        True,
        True,  # 16-20
        True,
        True,
        True,
        True,
        True,  # 21-25
        True,
        True,
        True,
        True,
        False,  # 26-30
        False,
        False,
        False,  # 31-33
    ]

    tmp_dir = Path("Salida") / "_tmp_fake_excel"
    tmp_dir.mkdir(exist_ok=True)
    replacements: Dict[str, Path] = {}

    # Week Dictionary
    week_headers = sheet_info["Week Dictionary"]["headers"]  # type: ignore
    week_rows = [[row["Time"], row["Week"]] for row in week_data]
    week_num_flags = [False, True]
    week_sheet_path = tmp_dir / "week.xml"
    write_table_sheet(week_sheet_path, week_headers, week_rows, week_num_flags)
    replacements[sheet_info["Week Dictionary"]["path"]] = week_sheet_path  # type: ignore

    # Brand Dictionary
    brand_headers = sheet_info["Brand Dictionary"]["headers"]  # type: ignore
    brand_rows_for_sheet = [[row["Brand"], row["Name"]] for row in brand_data]
    brand_sheet_path = tmp_dir / "brand.xml"
    write_table_sheet(brand_sheet_path, brand_headers, brand_rows_for_sheet, [False, False])
    replacements[sheet_info["Brand Dictionary"]["path"]] = brand_sheet_path  # type: ignore

    # Category Dictionary
    category_headers = sheet_info["Category Dictionary"]["headers"]  # type: ignore
    category_rows_for_sheet = [
        [row["Product"], row["Category"], row["Subcategory"]] for row in category_data
    ]
    category_sheet_path = tmp_dir / "category.xml"
    write_table_sheet(
        category_sheet_path, category_headers, category_rows_for_sheet, [False, False, False]
    )
    replacements[sheet_info["Category Dictionary"]["path"]] = category_sheet_path  # type: ignore

    # Source of Truth
    source_sheet_path = tmp_dir / "source.xml"
    write_source_sheet(
        source_sheet_path,
        source_headers,
        source_num_flags,
        source_row_count,
        brand_data,
        category_data,
        week_data,
    )
    replacements[sheet_info["Source of Truth"]["path"]] = source_sheet_path  # type: ignore

    with zipfile.ZipFile(src_path, "r") as src_zip, zipfile.ZipFile(
        dst_path, "w", compression=zipfile.ZIP_DEFLATED
    ) as out_zip:
        for item in src_zip.infolist():
            if item.filename in replacements:
                out_zip.write(replacements[item.filename], arcname=item.filename)
            else:
                out_zip.writestr(item, src_zip.read(item.filename))

    # Cleanup temporary XMLs
    for path in tmp_dir.iterdir():
        path.unlink()
    tmp_dir.rmdir()

    print(f"Archivo generado: {dst_path}")


if __name__ == "__main__":
    main()
