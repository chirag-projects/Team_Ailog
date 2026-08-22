"""Python API for the Unilog Content Enrichment Workbench.

Run: python server.py
Open: http://127.0.0.1:8000

The API accepts rows from any CSV/XLSX column layout. It detects common aliases,
cleans values, removes exact part-number duplicates, and returns Unilog delivery rows.
"""

from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Any

from flask import Flask, jsonify, request, send_from_directory


app = Flask(__name__, static_folder=".")

PLACEHOLDER = re.compile(r"^--\s*(unbranded|no .*brand)\s*--$", re.IGNORECASE)
OFFICIAL_DOMAINS = {
    r"diablo|freud": "diablotools.com",
    r"\b3m\b": "3m.com",
    r"bosch": "boschtools.com",
    r"dewalt": "dewalt.com",
    r"milwaukee": "milwaukeetool.com",
    r"whirlpool": "whirlpool.com",
    r"frigidaire": "frigidaire.com",
    r"rheem": "rheem.com",
}
BRAND_ALIASES = {
    "3m": "3M",
    "diablo tools": "Diablo",
    "dewalt": "DEWALT",
    "milwaukee tool": "Milwaukee Tool",
    "bosch tools": "Bosch",
}
BRAND_MASTER = {
    # Canonical Brand : (approved manufacturer, display brand with approved symbol)
    "3m": ("3M Company", "3M"),
    "diablo": ("Freud Inc.", "Diablo"),
    "frigidaire": ("Rheem Manufacturing", "FRIGIDAIRE®"),
    "whirlpool": ("Whirlpool Corporation", "Whirlpool®"),
    "dewalt": ("Stanley Black & Decker, Inc.", "DEWALT®"),
    "milwaukee": ("Milwaukee Tool", "Milwaukee®"),
    "bosch": ("Robert Bosch Tool Corporation", "Bosch®"),
}
TRADE_NAME_PATTERNS = [
    "Professional Series", "Eco Series", "Cubitron II", "Stikit", "CleanBoost™",
]
BLOCKED_COMMERCE_DOMAINS = [
    "amazon.com", "amazon.in", "flipkart.com", "ebay.com", "walmart.com",
    "aliexpress.com", "temu.com", "shopify.com",
]
UNICAT_MASTER: list[dict[str, str]] = []

TAXONOMY_RULES: list[dict[str, Any]] = [
    {
        "dept": "Appliances & Consumer Electronics",
        "class": "Kitchen Appliances",
        "fine": "Built-In Dishwashers",
        "classpath": "Appliances & Consumer Electronics>Kitchen Appliances>Built-In Dishwashers",
        "product_name": "Dishwasher",
        "lov": ["built-in dishwasher", "built in dishwasher", "dishwasher"],
    },
    {
        "dept": "Plumbing",
        "class": "Faucets",
        "fine": "Kitchen & Bath Sink Faucets",
        "classpath": "Plumbing>Faucets>Kitchen & Bath Sink Faucets",
        "product_name": "Faucet",
        "lov": ["faucet", "tap"],
    },
    {
        "dept": "Plumbing",
        "class": "Pipe Fittings",
        "fine": "Fittings",
        "classpath": "Plumbing>Pipe, Tube & Hose Fittings>Fittings",
        "product_name": "Fitting",
        "lov": ["coupling", "elbow", "tee", "fitting", "adapter"],
    },
    {
        "dept": "Abrasives",
        "class": "Bonded Abrasives",
        "fine": "Cut-Off Wheels",
        "classpath": "Abrasives>Bonded Abrasives>Cut-Off Wheels",
        "product_name": "Cut-Off Wheel",
        "lov": ["cut-off disc", "cutoff disc", "cut-off wheel", "cutoff wheel", "cutting wheel"],
    },
    {
        "dept": "Abrasives",
        "class": "Coated Abrasives",
        "fine": "Sanding Belts & Discs",
        "classpath": "Abrasives>Coated Abrasives>Sanding Belts & Discs",
        "product_name": "Sanding Disc",
        "lov": ["sanding belt", "sanding disc", "abrasive belt", "abrasive disc", "flap disc"],
    },
    {
        "dept": "Tools",
        "class": "Power Tool Accessories",
        "fine": "Cutting Tools",
        "classpath": "Tools>Power Tool Accessories>Cutting Tools",
        "product_name": "Power Tool Accessory",
        "lov": ["drill", "saw", "blade", "hole saw"],
    },
]

TAXONOMY_FALLBACK = ("Unclassified", "Unclassified", "Unclassified", "Unclassified", "Product")

# Category-specific allowed values (LOV) for categorical attributes.
FITTINGS_LOV = {
    "Connection Type": ["NPT", "BSP", "Slip", "Compression", "Push-to-Connect", "Barbed", "Flanged", "Socket Weld", "Threaded"],
    "Mounting Type": ["Wall Mount", "Floor Mount", "Inline"],
    "Material": ["Brass", "Carbon Steel", "Copper", "CPVC", "Galvanized Steel", "Nylon", "Plastic", "PVC", "Stainless Steel", "Steel"],
}

FAUCETS_LOV = {
    "Mounting Type": ["Deck Mount", "Wall Mount", "Single Hole", "Centerset", "Widespread"],
    "Handle Type": ["Single Handle", "Double Handle", "Touchless"],
    "Finish": ["Brushed Nickel", "Chrome", "Matte Black", "Oil Rubbed Bronze", "Polished Nickel", "Stainless"],
    "Material": ["Brass", "Stainless Steel", "Zinc"],
}

UNICAT_LOV = {
    "Material": ["Brass", "Carbon Steel", "Copper", "Nylon", "Plastic", "Stainless Steel", "Steel", "Zinc"],
    "Mounting Type": ["Built-In", "Deck Mount", "Inline", "Wall Mount"],
    "Power Source": ["Battery", "Corded Electric", "Electric", "Gas"],
}

CATEGORY_LOV_BY_FINE = {
    "Fittings": FITTINGS_LOV,
    "Kitchen & Bath Sink Faucets": FAUCETS_LOV,
    "Built-In Dishwashers": {
        "Mounting Type": ["Built-In"],
        "Power Source": ["Electric"],
        "Material": ["Plastic", "Stainless Steel", "Steel"],
    },
    "Sanding Belts & Discs": {
        "Backing Material": ["Cloth", "Fiber", "Paper"],
        "Attachment Type": ["Hook and Loop", "PSA"],
        "Material": ["Aluminum Oxide", "Ceramic", "Silicon Carbide", "Zirconia Alumina"],
    },
}


def create_schema() -> list[str]:
    columns = [
        "MFR URL", "Ref URL 1", "Ref URL 2", "Ref URL 3", "Ref URL 4", "Ref URL 5",
        "PART_NUMBER", "Dept", "Class", "Fine", "SKU - MY_PART_NUMBER",
        "Mfg_Part_Num", "Part_Desc", "E1_Brand", "Unilog_Brand", "DIB_Brand",
        "Part_Manuf", "MANUFACTURER_NAME", "BRAND_NAME", "TRADE_NAME",
        "MANUFACTURER_PART_NUMBER", "ALTERNATE_PART_NUMBER", "Classpath", "MOBILE_DESC",
        "INVOICE_DESC", "SHORT_DESC", "LONG_DESC1", "RETAIL_DESC", "MARKETING_DESCRIPTION",
    ]
    columns += [f"ITEM_FEATURES_{number}" for number in range(1, 21)]
    columns += ["With", "Standard/Approvals", "Prop 65", "Application", "Includes", "Product Name"]
    for number in range(1, 51):
        columns += [f"ATTRIBUTE_LABEL {number}", f"ATTRIBUTE_VALUE {number}", f"ATTRIBUTE_UOM {number}"]
    columns += [
        "UPC", "EAN", "GTIN", "UNSPSC", "Warranty", "List Price", "Selling Qty", "Selling UOM",
        "Standard Packaging Information", "LENGTH", "LENGTH_UOM", "HEIGHT", "HEIGHT_UOM", "WIDTH",
        "WIDTH_UOM", "WEIGHT", "WEIGHT_UOM", "VOLUME", "VOLUME_UOM", "Product Image",
        "Alternate Image 1", "Alternate Image 2", "Alternate Image 3", "Alternate Image 4", "SDS", "SDS_1",
        "Warranty Information", "Catalog", "Specification Sheet", "Instruction/Installation Manual",
        "Service Manual", "Owners/User Manual", "Line Drawing", "MTR", "RoHS", "Full Engineering Drawing",
        "Energy Star Guide", "Technical Bulletin", "Submittal", "Compatibility Chart", "Size Chart",
        "Product Label/Insert", "Video Link", "Video Link 1", "Country Of Origin", "Discontinued",
        "Actual Image (Yes/No)",
    ]
    return columns


SCHEMA = create_schema()


def canonical_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def read_value(row: dict[str, Any], aliases: list[str]) -> str:
    lookup = {canonical_header(key): value for key, value in row.items()}
    for alias in aliases:
        value = str(lookup.get(canonical_header(alias), "")).strip()
        if value and not PLACEHOLDER.fullmatch(value):
            return value
    return ""


def comparable(value: str) -> str:
    """Normalize text for entity resolution while keeping the original official value."""
    return re.sub(r"[^a-z0-9]", "", value.lower())


def load_unicat_master(rows: list[dict[str, Any]]) -> int:
    """Store official manufacturer/brand pairs from a UniCat workbook export."""
    UNICAT_MASTER.clear()
    seen: set[tuple[str, str]] = set()
    for row in rows:
        manufacturer = read_value(row, ["MANUFACTURER_NAME", "manufacturer name"])
        brand = read_value(row, ["BRAND_NAME", "brand name"])
        if not manufacturer or not brand:
            continue
        key = (comparable(manufacturer), comparable(brand))
        if key in seen:
            continue
        seen.add(key)
        UNICAT_MASTER.append({"manufacturer": manufacturer, "brand": brand})
    return len(UNICAT_MASTER)


def fuzzy_unicat_match(raw_brand: str, description: str, raw_manufacturer: str) -> dict[str, str] | None:
    """Return the best official UniCat pair, but only when the match is reliable.

    Brand is weighted most heavily because distributor/manufacturer fields commonly
    contain supplier account names rather than the actual product manufacturer.
    """
    if not UNICAT_MASTER:
        return None
    search_brand = comparable(raw_brand)
    search_manufacturer = comparable(clean_manufacturer_name(raw_manufacturer))
    search_description = comparable(description)
    best_match: dict[str, str] | None = None
    best_score = 0.0
    for entry in UNICAT_MASTER:
        official_brand = comparable(entry["brand"])
        official_manufacturer = comparable(entry["manufacturer"])
        brand_score = SequenceMatcher(None, search_brand, official_brand).ratio() if search_brand else 0.0
        manufacturer_score = SequenceMatcher(None, search_manufacturer, official_manufacturer).ratio() if search_manufacturer else 0.0
        description_score = 1.0 if official_brand and official_brand in search_description else 0.0
        score = max(brand_score * 0.75 + manufacturer_score * 0.25, description_score * 0.9 + manufacturer_score * 0.1)
        if score > best_score:
            best_score = score
            best_match = entry
    return best_match if best_score >= 0.78 else None


def clean_manufacturer_name(value: str) -> str:
    """Remove supplier account codes, keeping the readable legal company name."""
    clean = re.sub(r"\s*\([^)]*\)\s*$", "", value).strip()
    clean = re.sub(r"\s+", " ", clean)
    known_names = {
        "freud inc": "Freud Inc.",
        "freud inc.": "Freud Inc.",
        "whirlpool corporation": "Whirlpool Corporation",
        "rheem manufacturing": "Rheem Manufacturing",
        "3m company": "3M Company",
    }
    return known_names.get(clean.lower(), clean)


def resolve_brand_and_manufacturer(raw_brand: str, description: str, raw_manufacturer: str) -> tuple[str, str]:
    """Use submitted brand first, then infer a known brand from the product description."""
    master_match = fuzzy_unicat_match(raw_brand, description, raw_manufacturer)
    if master_match:
        return master_match["brand"], master_match["manufacturer"]

    brand_candidate = BRAND_ALIASES.get(raw_brand.lower(), raw_brand).strip()
    searchable_text = f"{brand_candidate} {description}".lower()

    for brand_key, (manufacturer, display_brand) in BRAND_MASTER.items():
        if re.search(rf"\b{re.escape(brand_key)}\b", searchable_text, re.IGNORECASE):
            return display_brand, manufacturer

    return brand_candidate, clean_manufacturer_name(raw_manufacturer)


def extract_trade_name(description: str) -> str:
    """Capture a named series or trade line only when it is present in the source text."""
    for trade_name in TRADE_NAME_PATTERNS:
        if trade_name.lower() in description.lower():
            return trade_name
    return ""


def decimal_to_fraction(value: str) -> str:
    number = float(value)
    whole = int(number)
    numerator = round((number - whole) * 64)
    if numerator == 0:
        return str(whole)
    if numerator == 64:
        return str(whole + 1)
    denominator = 64
    while numerator % 2 == 0 and denominator % 2 == 0:
        numerator //= 2
        denominator //= 2
    return f"{whole}-{numerator}/{denominator}" if whole else f"{numerator}/{denominator}"


def normalize_units(value: str) -> str:
    text = re.sub(r"\s+", " ", value).strip()
    replacements = [
        (r"(\d)\s*(?:inches|inch|ins?\.?)(?=\s|$|x|×|,|\))", r"\1 in"),
        (r"(\d)\s*(?:volts?|v\.)(?=\s|$|,|\))", r"\1 V"),
        (r"(\d)\s*(?:amps?|amperes?|a\.)(?=\s|$|,|\))", r"\1 A"),
        (r"(\d)\s*(?:pounds?|lbs?\.?)(?=\s|$|,|\))", r"\1 lb"),
    ]
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return re.sub(r"\b(\d+\.\d+)\s*in\b", lambda match: f"{decimal_to_fraction(match.group(1))} in", text, flags=re.IGNORECASE)


def classify(description: str) -> tuple[str, str, str, str, str]:
    """Assign taxonomy using LOV phrase matching and rule disambiguation."""
    text = re.sub(r"\s+", " ", description.lower()).strip()
    if not text:
        return TAXONOMY_FALLBACK

    for rule in TAXONOMY_RULES:
        for phrase in sorted(rule["lov"], key=len, reverse=True):
            if re.search(rf"\b{re.escape(phrase)}\b", text):
                product_name = rule["product_name"]
                if rule["fine"] == "Sanding Belts & Discs":
                    if "belt" in text:
                        product_name = "Sanding Belt"
                    elif "disc" in text or "flap" in text:
                        product_name = "Sanding Disc"
                return rule["dept"], rule["class"], rule["fine"], rule["classpath"], product_name

    if re.search(r"\bdisc\b", text) and re.search(r"cut\s*-?\s*off|cutting", text):
        return "Abrasives", "Bonded Abrasives", "Cut-Off Wheels", "Abrasives>Bonded Abrasives>Cut-Off Wheels", "Cut-Off Wheel"
    if re.search(r"abrasiv|sand", text):
        name = "Sanding Belt" if "belt" in text else "Sanding Disc"
        return "Abrasives", "Coated Abrasives", "Sanding Belts & Discs", "Abrasives>Coated Abrasives>Sanding Belts & Discs", name

    return TAXONOMY_FALLBACK


def allowed_values_for(fine: str, label: str) -> list[str]:
    category_lov = CATEGORY_LOV_BY_FINE.get(fine, {})
    if label in category_lov:
        return category_lov[label]
    return UNICAT_LOV.get(label, [])


def lov_match(label: str, text: str, fine: str) -> str:
    options = allowed_values_for(fine, label)
    for option in sorted(options, key=len, reverse=True):
        if re.search(rf"\b{re.escape(option)}\b", text, re.IGNORECASE):
            return option
    return ""


def extract_series(description: str) -> str:
    for trade_name in TRADE_NAME_PATTERNS:
        if trade_name.lower() in description.lower():
            return trade_name
    direct = re.search(r"\b([A-Za-z0-9+&\- ]{2,40}\sSeries)\b", description, re.IGNORECASE)
    return direct.group(1) if direct else ""


def extract_attributes(description: str, fine: str) -> list[dict[str, str]]:
    text = normalize_units(description)
    attributes: list[dict[str, str]] = []
    seen_labels: set[str] = set()

    def add(label: str, value: str, uom: str = "") -> None:
        if not value or label in seen_labels:
            return
        options = allowed_values_for(fine, label)
        canonical = value
        if options:
            canonical = next((option for option in options if comparable(option) == comparable(value)), "")
            if not canonical:
                return
        attributes.append({"label": label, "value": canonical, "uom": uom})
        seen_labels.add(label)

    series = extract_series(text)
    if series:
        add("Series", series)

    size = re.search(r"(\d+(?:-\d+/\d+|/\d+|\.\d+)?)(?:\s*(?:in|\"))?\s*[x×]\s*(\d+(?:-\d+/\d+|/\d+|\.\d+)?)(?:\s*(?:in|\"))(?:\s*[x×]\s*(\d+(?:-\d+/\d+|/\d+|\.\d+)?)(?:\s*(?:in|\"))?)?", text, re.IGNORECASE)
    if size:
        dims = [size.group(1), size.group(2)]
        if size.group(3):
            dims.append(size.group(3))
        add("Size", " x ".join(f"{dimension} in" for dimension in dims))

    connection_size = re.search(r"\b(\d+(?:-\d+/\d+|/\d+|\.\d+)?)\s*(?:in|\")\b", text, re.IGNORECASE)
    if connection_size and fine == "Fittings":
        add("Connection Size", connection_size.group(1), "in")

    if fine in {"Sanding Belts & Discs", "Cut-Off Wheels"}:
        grit_match = re.search(r"\b(?:grit\s*[:\-]?\s*(P?\d{2,4})|(P?\d{2,4})\s*grit)\b", text, re.IGNORECASE)
        if grit_match:
            grit_value = grit_match.group(1) or grit_match.group(2) or ""
            add("Grit", grit_value.upper().replace(" ", ""))

    quantity = re.search(r"\b(\d+)\s*(pc|pcs|piece|pieces|pack|set)\b", text, re.IGNORECASE)
    if quantity:
        add("Package Quantity", quantity.group(1), "EA")

    material = lov_match("Material", text, fine)
    if material:
        add("Material", material)

    mounting_type = lov_match("Mounting Type", text, fine)
    if mounting_type:
        add("Mounting Type", mounting_type)

    connection_type = lov_match("Connection Type", text, fine)
    if connection_type:
        add("Connection Type", connection_type)

    handle_type = lov_match("Handle Type", text, fine)
    if handle_type:
        add("Handle Type", handle_type)

    finish = lov_match("Finish", text, fine)
    if finish:
        add("Finish", finish)

    backing = lov_match("Backing Material", text, fine)
    if backing:
        add("Backing Material", backing)

    attachment = lov_match("Attachment Type", text, fine)
    if attachment:
        add("Attachment Type", attachment)

    power_source = lov_match("Power Source", text, fine)
    if power_source:
        add("Power Source", power_source)

    voltage = re.search(r"\b(\d+(?:\.\d+)?)\s*V\b", text, re.IGNORECASE)
    if voltage:
        add("Voltage Rating", voltage.group(1), "V")

    amperage = re.search(r"\b(\d+(?:\.\d+)?)\s*A\b", text, re.IGNORECASE)
    if amperage:
        add("Amperage Rating", amperage.group(1), "A")

    flow_rate = re.search(r"\b(\d+(?:\.\d+)?)\s*gpm\b", text, re.IGNORECASE)
    if flow_rate:
        add("Flow Rate", flow_rate.group(1), "GPM")

    sound_level = re.search(r"\b(\d+(?:\.\d+)?)\s*dB\b", text, re.IGNORECASE)
    if sound_level:
        add("Sound Level", sound_level.group(1), "dB")

    handles = re.search(r"\b(\d)\s*handle\b", text, re.IGNORECASE)
    if handles:
        add("Number of Handles", handles.group(1))

    if fine == "Built-In Dishwashers":
        add("Mounting Type", "Built-In")
        add("Power Source", "Electric")

    width = re.search(r"\b(\d+(?:\.\d+)?)\s*in\s*(?:wide|width)\b", text, re.IGNORECASE)
    if width:
        add("Width", width.group(1), "in")

    height = re.search(r"\b(\d+(?:\.\d+)?)\s*in\s*(?:high|height)\b", text, re.IGNORECASE)
    if height:
        add("Height", height.group(1), "in")

    length = re.search(r"\b(\d+(?:\.\d+)?)\s*in\s*(?:long|length)\b", text, re.IGNORECASE)
    if length:
        add("Length", length.group(1), "in")

    return attributes


def attribute_value(attributes: list[dict[str, str]], label: str) -> str:
    for attribute in attributes:
        if attribute["label"] == label:
            value = attribute["value"]
            uom = attribute["uom"]
            return f"{value} {uom}".strip()
    return ""


def compress_size_for_short(size_value: str) -> str:
    if not size_value:
        return ""
    compact = re.sub(r"\s*in\b", '"', size_value, flags=re.IGNORECASE)
    compact = re.sub(r"\s*[x×]\s*", " x ", compact)
    return compact


def build_key_specs(attributes: list[dict[str, str]]) -> list[str]:
    specs: list[str] = []
    size = compress_size_for_short(attribute_value(attributes, "Size"))
    grit = attribute_value(attributes, "Grit")
    package_quantity = attribute_value(attributes, "Package Quantity")
    material = attribute_value(attributes, "Material")
    voltage = attribute_value(attributes, "Voltage Rating")
    amperage = attribute_value(attributes, "Amperage Rating")

    if size:
        specs.append(size)
    if grit:
        specs.append(f"{grit} Grit")
    if package_quantity:
        specs.append(package_quantity)
    if material:
        specs.append(material)
    if voltage:
        specs.append(f"{voltage}V")
    if amperage:
        specs.append(f"{amperage}A")
    return specs


def build_mobile_desc(brand: str, manufacturer: str, product_name: str, trade_name: str, part: str) -> str:
    display_brand = brand or manufacturer
    return ", ".join(item for item in [display_brand, product_name, trade_name, part] if item)


def build_invoice_desc(product_name: str, key_specs: list[str], part: str) -> str:
    for count in range(len(key_specs), -1, -1):
        candidate = " ".join(item for item in [product_name, " ".join(key_specs[:count]), part] if item).upper()
        if len(candidate) <= 40:
            return candidate
    fallback = f"{product_name} {part}".upper()
    return fallback[:40]


def build_short_desc(brand: str, trade_name: str, part: str, key_specs: list[str], product_name: str) -> str:
    prefix = [brand, trade_name, part]
    detail = " ".join(key_specs[:2]) if key_specs else product_name
    return " ".join(item for item in [*prefix, detail] if item)


def build_long_desc1(brand: str, manufacturer: str, trade_name: str, product_name: str, part: str, key_specs: list[str], description: str) -> str:
    header = " ".join(item for item in [brand or manufacturer, trade_name, product_name] if item)
    details = "; ".join(key_specs)
    if details:
        return f"{header} ({part}). Key specs: {details}. Source description: {description}."
    return f"{header} ({part}). Source description: {description}."


def source_urls(brand: str, manufacturer: str, part: str) -> dict[str, str]:
    matching_domain = next((domain for pattern, domain in OFFICIAL_DOMAINS.items() if re.search(pattern, f"{brand} {manufacturer}", re.IGNORECASE)), "")
    if matching_domain:
        return {"domain": matching_domain, "search": f"https://www.google.com/search?q=site%3A{matching_domain}+%22{part}%22", "home": f"https://{matching_domain}", "reference_search": ""}
    excluded = "+".join(f"-site%3A{domain}" for domain in BLOCKED_COMMERCE_DOMAINS)
    return {"domain": "", "search": "", "home": "", "reference_search": f"https://www.google.com/search?q=%22{part}%22+%22{brand or manufacturer}%22+{excluded}"}


def enrich_row(row: dict[str, Any]) -> dict[str, Any]:
    part = read_value(row, ["Mfg_Part_Num", "manufacturer part number", "mpn", "part number", "sku"]) or "UNASSIGNED"
    raw_description = read_value(row, ["Part_Desc", "description", "product description", "name"]) or part
    description = normalize_units(raw_description)
    raw_brand = read_value(row, ["Unilog_Brand", "E1_Brand", "DIB_Brand", "brand"])
    raw_manufacturer = read_value(row, ["Part_Manuf", "manufacturer", "mfr"])
    brand, manufacturer = resolve_brand_and_manufacturer(raw_brand, description, raw_manufacturer)
    manufacturer = manufacturer or brand or "Unknown Manufacturer"
    trade_name = extract_trade_name(description)
    dept, product_class, fine, classpath, product_name = classify(description)
    attributes = extract_attributes(description, fine)
    source = source_urls(brand, manufacturer, part)
    record = dict.fromkeys(SCHEMA, "")
    key_specs = build_key_specs(attributes)
    short = build_short_desc(brand, trade_name, part, key_specs, product_name)
    long_desc = build_long_desc1(brand, manufacturer, trade_name, product_name, part, key_specs, description)
    mobile_desc = build_mobile_desc(brand, manufacturer, product_name, trade_name, part)
    invoice_desc = build_invoice_desc(product_name, key_specs, part)
    record.update({
        "MFR URL": source["search"], "Ref URL 1": source["home"], "Ref URL 2": source["reference_search"],
        "PART_NUMBER": read_value(row, ["PART_NUMBER", "part number"]), "SKU - MY_PART_NUMBER": read_value(row, ["SKU - MY_PART_NUMBER", "sku"]),
        "Mfg_Part_Num": part, "Part_Desc": description, "E1_Brand": read_value(row, ["E1_Brand"]), "Unilog_Brand": read_value(row, ["Unilog_Brand"]), "DIB_Brand": read_value(row, ["DIB_Brand"]), "Part_Manuf": read_value(row, ["Part_Manuf"]),
        "MANUFACTURER_NAME": manufacturer, "BRAND_NAME": brand, "TRADE_NAME": trade_name, "MANUFACTURER_PART_NUMBER": part,
        "Dept": dept, "Class": product_class, "Fine": fine, "Classpath": classpath, "Product Name": product_name,
        "SHORT_DESC": short, "LONG_DESC1": long_desc,
        "RETAIL_DESC": ", ".join(item for item in [product_name, ", ".join(attribute["label"] for attribute in attributes)] if item),
        "MOBILE_DESC": mobile_desc,
        "INVOICE_DESC": invoice_desc, "UPC": read_value(row, ["UPC"]), "EAN": read_value(row, ["EAN"]), "GTIN": read_value(row, ["GTIN"]), "UNSPSC": read_value(row, ["UNSPSC"]),
    })
    for index, attribute in enumerate(attributes, start=1):
        record[f"ATTRIBUTE_LABEL {index}"] = attribute["label"]
        record[f"ATTRIBUTE_VALUE {index}"] = attribute["value"]
        record[f"ATTRIBUTE_UOM {index}"] = attribute["uom"]
    reasons = [reason for reason in [
        "Missing manufacturer part number" if part == "UNASSIGNED" else "",
        "Unclassified product type" if classpath == "Unclassified" else "",
        "Brand unresolved" if not brand else "",
        "Official manufacturer domain not mapped" if not source["domain"] else "",
    ] if reason]
    score = min(100, (25 if part != "UNASSIGNED" else 0) + 25 + (15 if brand else 0) + (20 if classpath != "Unclassified" else 0) + (10 if source["domain"] else 0) + min(5, len(attributes) * 2))
    record["_review"] = {"attributes": attributes, "source": source, "reasons": reasons, "confidence": score}
    return record


def clean_and_deduplicate(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    unique: dict[str, dict[str, Any]] = {}
    duplicates = 0
    for raw_row in rows:
        row = {str(key).strip(): str(value or "").strip() for key, value in raw_row.items()}
        part = read_value(row, ["Mfg_Part_Num", "manufacturer part number", "mpn", "part number", "sku"])
        description = read_value(row, ["Part_Desc", "description", "product description", "name"])
        key = re.sub(r"[^A-Z0-9]", "", (part or description).upper())
        if not key: continue
        if key not in unique:
            unique[key] = row
            continue
        duplicates += 1
        for column, value in row.items():
            if not unique[key].get(column) and value:
                unique[key][column] = value
    return list(unique.values()), duplicates


@app.get("/")
def index() -> Any:
    return send_from_directory(".", "index.html")


@app.post("/api/enrich")
def enrich_api() -> Any:
    body = request.get_json(silent=True) or {}
    rows = body.get("rows", [])
    if not isinstance(rows, list):
        return jsonify({"error": "Expected a JSON array of rows."}), 400
    cleaned_rows, duplicates = clean_and_deduplicate(rows)
    return jsonify({"records": [enrich_row(row) for row in cleaned_rows], "duplicates": duplicates, "schemaLength": len(SCHEMA)})


@app.post("/api/unicat-master")
def unicat_master_api() -> Any:
    body = request.get_json(silent=True) or {}
    rows = body.get("rows", [])
    if not isinstance(rows, list):
        return jsonify({"error": "Expected a JSON array of UniCat rows."}), 400
    count = load_unicat_master(rows)
    if not count:
        return jsonify({"error": "Could not find MANUFACTURER_NAME and BRAND_NAME columns in this file."}), 400
    return jsonify({"loaded": count})


@app.get("/<path:filename>")
def static_files(filename: str) -> Any:
    return send_from_directory(".", filename)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8000, debug=True)
