#!/usr/bin/env python3
"""
Extract MDSA premium data from sharedStrings.xml into structured JSON.

Parses the Excel shared strings XML to find premium tables for each
gender/age combination and outputs premium_data.json.
"""

import json
import re
import xml.etree.ElementTree as ET


def extract_strings(xml_path):
    """Parse sharedStrings.xml and return list of all shared string values."""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    strings = []
    for si in root.findall(f"{{{ns}}}si"):
        texts = []
        for t in si.iter(f"{{{ns}}}t"):
            if t.text:
                texts.append(t.text)
        strings.append("".join(texts))

    return strings


def is_data_string(s, min_tokens=5):
    """
    Check if a string is a numeric data string.
    Data strings have tokens that are numbers (possibly with commas/decimals) or '-'.
    We require at least min_tokens numeric-like tokens at the start.
    """
    tokens = s.replace("\n", " ").strip().split()
    if len(tokens) < min_tokens:
        return False

    numeric_count = 0
    for token in tokens[:min_tokens]:
        cleaned = token.replace(",", "")
        if token == "-":
            numeric_count += 1
        elif re.match(r"^-?\d+\.?\d*$", cleaned):
            numeric_count += 1
        else:
            break

    return numeric_count >= min_tokens


def parse_idr_values(s):
    """
    Parse IDR data string into list of 44 values.
    Values are in Juta Rupiah. '-' means not available (null).
    The string has 44 numeric values followed by trailing text like 'Premi Dasar Tahunan'.
    """
    tokens = s.replace("\n", " ").strip().split()
    values = []
    for token in tokens:
        if token == "-":
            values.append(None)
        elif re.match(r"^-?\d+\.?\d*$", token):
            values.append(float(token))
        else:
            # Stop at first non-numeric token (header text at end)
            break

    return values[:44]


def parse_usd_values(s):
    """
    Parse USD data string into list of 18 values.
    Values may have commas (e.g., '3,160'). '-' means not available (null).
    """
    tokens = s.replace("\n", " ").strip().split()
    values = []
    for token in tokens:
        if token == "-":
            values.append(None)
        else:
            cleaned = token.replace(",", "")
            if re.match(r"^-?\d+\.?\d*$", cleaned):
                val = float(cleaned) if "." in cleaned else int(cleaned)
                values.append(val)
            else:
                break

    return values[:18]


def main():
    xml_path = "/projects/sandbox/Project-mdsa/xlsx_extract/xl/sharedStrings.xml"
    output_path = "/projects/sandbox/Project-mdsa/premium_data.json"

    strings = extract_strings(xml_path)
    print(f"Total shared strings: {len(strings)}")

    # Pattern to match title strings like "Pria - Usia 0 Tahun" or "Wanita - Usia 70 Tahun"
    title_pattern = re.compile(
        r"(Pria|Wanita)\s*[\u2013-]\s*Usia\s+(\d+)\s+Tahun"
    )

    # Find all title indices
    titles = []
    for i, s in enumerate(strings):
        m = title_pattern.search(s)
        if m:
            gender = m.group(1).lower()
            age = int(m.group(2))
            titles.append((i, gender, age))

    print(f"Found {len(titles)} title entries")

    # For each title, find the next IDR data string and the next USD data string
    premium_data = {}

    for idx, (title_idx, gender, age) in enumerate(titles):
        key = f"{gender}|{age}"

        # Determine search boundary (up to next title or end of strings)
        if idx + 1 < len(titles):
            boundary = titles[idx + 1][0]
        else:
            boundary = len(strings)

        # Find first data string after title (IDR - expect ~44-47 tokens)
        idr_idx = None
        for i in range(title_idx + 1, boundary):
            if is_data_string(strings[i], min_tokens=5):
                tokens = strings[i].replace("\n", " ").strip().split()
                # IDR strings have 44+ values (44 data + possible trailing text)
                numeric_count = 0
                for t in tokens:
                    if t == "-" or re.match(r"^-?\d+\.?\d*$", t):
                        numeric_count += 1
                    else:
                        break
                if numeric_count >= 18:
                    # This is the IDR string (more than 18 values)
                    idr_idx = i
                    break

        if idr_idx is None:
            print(f"WARNING: No IDR data found for {key} (title at {title_idx})")
            continue

        # Find next data string after IDR (USD - expect exactly 18 tokens)
        usd_idx = None
        for i in range(idr_idx + 1, boundary):
            if is_data_string(strings[i], min_tokens=5):
                tokens = strings[i].replace("\n", " ").strip().split()
                # USD strings have around 18 values
                numeric_count = 0
                for t in tokens:
                    cleaned = t.replace(",", "")
                    if t == "-" or re.match(r"^-?\d+\.?\d*$", cleaned):
                        numeric_count += 1
                    else:
                        break
                if numeric_count >= 10:
                    usd_idx = i
                    break

        if usd_idx is None:
            print(f"WARNING: No USD data found for {key} (title at {title_idx}, IDR at {idr_idx})")
            # Still include with empty USD
            idr_values = parse_idr_values(strings[idr_idx])
            premium_data[key] = {"idr": idr_values, "usd": []}
            continue

        idr_values = parse_idr_values(strings[idr_idx])
        usd_values = parse_usd_values(strings[usd_idx])

        premium_data[key] = {"idr": idr_values, "usd": usd_values}

    # Save output
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(premium_data, f, indent=2, ensure_ascii=False)

    print(f"\nSaved {len(premium_data)} entries to {output_path}")

    # Verification
    print("\n--- Verification ---")
    if "pria|0" in premium_data:
        idr = premium_data["pria|0"]["idr"]
        usd = premium_data["pria|0"]["usd"]
        print(f"pria|0 IDR ({len(idr)} values): {idr[:10]}")
        print(f"pria|0 USD ({len(usd)} values): {usd[:5]}")

    if "pria|61" in premium_data:
        idr = premium_data["pria|61"]["idr"]
        print(f"pria|61 IDR ({len(idr)} values): {idr[:8]}")

    if "wanita|70" in premium_data:
        idr = premium_data["wanita|70"]["idr"]
        usd = premium_data["wanita|70"]["usd"]
        print(f"wanita|70 IDR ({len(idr)} values): {idr[:5]}")
        print(f"wanita|70 USD ({len(usd)} values): {usd[:5]}")

    # Check all entries have correct counts
    issues = []
    for key, data in premium_data.items():
        if len(data["idr"]) != 44:
            issues.append(f"{key}: IDR has {len(data['idr'])} values (expected 44)")
        if len(data["usd"]) != 18:
            issues.append(f"{key}: USD has {len(data['usd'])} values (expected 18)")

    if issues:
        print(f"\nISSUES ({len(issues)}):")
        for issue in issues[:20]:
            print(f"  {issue}")
    else:
        print("\nAll entries have correct value counts (44 IDR, 18 USD)")


if __name__ == "__main__":
    main()
