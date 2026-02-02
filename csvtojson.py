#!/usr/bin/env python3
"""
csvtojson_force_switch1_no_fill.py

Parses a spreadsheet-style CSV containing repeated 5-column blocks:
  Port,vLAN,Dj,Alias,Room

Produces a JSON array of objects with the schema:
  {
    "FIELD1": "<port>",
    "Switch 1": "<vlan or blank>",
    "FIELD3": "<dj or blank>",
    "FIELD4": "<alias or blank>",
    "FIELD5": "<room or blank>"
  }

Rules:
- Port must be digits and slashes (e.g. 1/1/1). Invalid ports are skipped.
- VLAN is preserved only if it's digits or 'lagg' (case-insensitive). Otherwise it's blank.
- Do NOT fill-down or back-fill VLAN values from merged cells.
- If VLAN is blank for a port, then FIELD3/FIELD4/FIELD5 are set to "" for that port.
- All VLAN values are written under the fixed key "Switch 1".
- Output sorted numerically by port.
"""
import csv
import json
import argparse
import re

PORT_RE = re.compile(r"^\d+(?:/\d+)+$")   # e.g. 1/1/1
VLAN_DIGIT_RE = re.compile(r"^\d+$")

def norm(s: str) -> str:
    """Trim, remove NBSP, preserve empties."""
    return (s or "").replace("\u00A0", " ").strip()

def clean_port(s: str) -> str:
    """Normalize port cell (remove leading ' used by Excel, trim)."""
    s = norm(s)
    if s.startswith("'"):
        s = s[1:].strip()
    return s

def clean_vlan_raw(s: str) -> str:
    """Normalize raw VLAN cell (trim, remove leading ', convert 103.0 -> 103)."""
    s = norm(s)
    if s.startswith("'"):
        s = s[1:].strip()
    m = re.fullmatch(r"(\d+)(?:\.0+)?", s)
    if m:
        return m.group(1)
    return s

def normalize_vlan_for_output(s: str) -> str:
    """
    Return digits or 'lagg' or blank. Anything else -> blank.
    This function does NOT attempt to fill blank values from neighboring rows.
    """
    s = clean_vlan_raw(s)
    if s == "":
        return ""
    if s.lower() == "lagg":
        return "lagg"
    if VLAN_DIGIT_RE.match(s):
        return s
    return ""  # invalid -> blank

def is_empty_row(row) -> bool:
    return not row or all(norm(c) == "" for c in row)

def find_block_starts(row):
    """
    Detect repeated "Port,vLAN,Dj,Alias,Room" groups across the row.
    Returns list of starting indices.
    """
    starts = []
    lower = [norm(c).lower() for c in row]
    for i in range(len(lower) - 4):
        if lower[i:i+5] == ["port", "vlan", "dj", "alias", "room"]:
            starts.append(i)
    return starts

def detect_switch_labels(row_above, block_starts):
    """
    For each block start, attempt to find a surrounding cell in the row above
    saying "Switch ..." — otherwise fallback to "Switch N".
    We still ignore these labels at output time (we force "Switch 1"),
    but we detect them to correctly align blocks when parsing multiple sections.
    """
    labels = {}
    for idx, start in enumerate(block_starts, 1):
        label = ""
        if row_above:
            for j in range(start, start + 5):
                if j < len(row_above):
                    cell = norm(row_above[j])
                    if cell.lower().startswith("switch"):
                        label = cell
                        break
        labels[start] = label if label else f"Switch {idx}"
    return labels

def valid_port(port: str) -> bool:
    """Port must match digits and slashes e.g. 1/1/1"""
    return bool(PORT_RE.match(port))

def port_to_sort_tuple(port: str):
    """Numeric tuple for sorting e.g. '3/1/46' -> (3,1,46)"""
    try:
        return tuple(int(x) for x in port.split("/"))
    except Exception:
        return (10**9,)

def get_cell(row, idx) -> str:
    return norm(row[idx]) if idx < len(row) else ""

def main():
    parser = argparse.ArgumentParser(description="Parse multi-switch CSV into normalized JSON (force Switch 1, no VLAN fill).")
    parser.add_argument("csvfile", help="Input CSV file")
    parser.add_argument("jsonfile", nargs="?", default="output.json", help="Output JSON file")
    args = parser.parse_args()

    with open(args.csvfile, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))

    if not rows:
        raise SystemExit("CSV is empty.")

    # Find the first header row that has repeated Port,vLAN,Dj,Alias,Room groups
    header_idx = None
    block_starts = None
    for i, row in enumerate(rows):
        starts = find_block_starts(row)
        if starts:
            header_idx = i
            block_starts = starts
            break

    if header_idx is None:
        raise SystemExit('Could not find a header row containing "Port,vLAN,Dj,Alias,Room".')

    # Initial switch labels (for parsing alignment only)
    row_above = rows[header_idx - 1] if header_idx > 0 else None
    switch_label_by_start = detect_switch_labels(row_above, block_starts)

    records = []

    # Parse rows after header; allow handling of later Switch 3/4 sections by detecting headers again
    i = header_idx + 1
    while i < len(rows):
        row = rows[i]
        if is_empty_row(row):
            i += 1
            continue

        # If this row is another header row, update block_starts and labels and skip it
        new_starts = find_block_starts(row)
        if new_starts:
            # update detected switch labels from the row above this header (if available)
            row_above_this = rows[i - 1] if i - 1 >= 0 else None
            switch_label_by_start.update(detect_switch_labels(row_above_this, new_starts))
            # update block_starts to include new ones (so subsequent data rows get parsed)
            for ns in new_starts:
                if ns not in block_starts:
                    block_starts.append(ns)
            i += 1
            continue

        # For each known block start, extract fields
        for start in sorted(block_starts):
            port_cell = get_cell(row, start)
            vlan_cell = get_cell(row, start + 1)
            dj_cell = get_cell(row, start + 2)
            room_cell = get_cell(row, start + 4)

            # If entire block is empty on this row, skip it
            if port_cell == "" and vlan_cell == "" and dj_cell == "" and room_cell == "":
                continue

            port = clean_port(port_cell)
            if not valid_port(port):
                # skip non-port rows (headers / garbage)
                continue

            # Normalize VLAN but DO NOT fill from other rows; keep blank if invalid or absent
            vlan_norm = normalize_vlan_for_output(vlan_cell)

            # If VLAN is blank, the rest should be blank for that port as well per your request
            if vlan_norm == "":
                dj_out = ""
                room_out = ""
                alias_out = ""
            else:
                dj_out = dj_cell
                room_out = room_cell
                if room_out and dj_out:
                    alias_out = f"{room_out} {dj_out}"
                elif room_out:
                    alias_out = room_out
                else:
                    alias_out = dj_out

            # Force VLAN under "Switch 1" key
            rec = {
                "FIELD1": port,
                "Switch 1": vlan_norm,
                "FIELD3": dj_out,
                "FIELD4": alias_out,
                "FIELD5": room_out
            }
            records.append(rec)

        i += 1

    # Sort numerically by port
    records.sort(key=lambda r: port_to_sort_tuple(r["FIELD1"]))

    # Write JSON array
    with open(args.jsonfile, "w", encoding="utf-8") as out_f:
        json.dump(records, out_f, indent=2)

    print(f"Done. Wrote {len(records)} records to {args.jsonfile}")

if __name__ == "__main__":
    main()
