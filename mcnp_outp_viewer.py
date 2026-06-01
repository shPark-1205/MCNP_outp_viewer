#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MCNP outp interactive viewer (PySide6/Qt) - v13 sortable tally list and Excel-number input

Features
--------
- Open one or more MCNP outp text files.
- Show tally list on the left: file, tally number, tally type, FC name.
- Click a tally to extract and display:
    1) Simple one-value tally such as TBR: value + relative error
    2) F6 + FS segmented tally: z_start, z_end, z_mid, value, relative error
    3) Energy-bin tally such as F1/F2/F4 + E card: bin, energy lower/upper, delta_E, log_delta_E, raw tally, tally/delta_E, tally/log_delta_E, relative error
- Export the currently displayed table to CSV.
- Check multiple tallies on the left and export all checked tallies to one XLSX workbook.
- The first XLSX sheet stacks all checked tallies; following sheets store individual tally tables.
- Check all currently visible tallies after filtering.
- Apply user-defined multiplier and additive offset to extracted tally values while preserving original values.

Install
-------
pip install PySide6 pandas openpyxl

Run
---
python mcnp_outp_viewer.py

Notes
-----
This parser is intentionally text-pattern based because MCNP outp files are plain text.
It works best when the outp contains the echoed input cards, which is the default MCNP style.
"""

from __future__ import annotations

import csv
import math
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

FLOAT_RE = r"[-+]?(?:\d+\.\d*|\.\d+|\d+)(?:[Ee][+-]?\d+)?"
APP_VERSION = "2026-05-28-v13-sort-excel-number"
UI_FONT_POINT_SIZE = 11.0
PREVIEW_ROW_LIMIT_DEFAULT = 500


def parse_user_number(text_value, default: float = 0.0) -> float:
    """Parse a user-entered number, including common Excel paste formats.

    Supported examples:
      1,234.56     -> 1234.56
      1.23E+4      -> 12300.0
      1.123.E+2    -> 112.3
      1.E-3        -> 0.001
    """
    if text_value is None:
        return default

    s = str(text_value).strip()
    if not s:
        return default

    # Excel or locale-style copy/paste sometimes includes thousands separators/spaces.
    s = s.replace(",", "").replace(" ", "").replace("\t", "")

    # Excel-like scientific string occasionally appears as 1.123.E+2.
    # Remove the dot immediately before E/e when it follows a digit.
    s = re.sub(r"(?<=\d)\.(?=[eE][+-]?\d+$)", "", s)

    try:
        return float(s)
    except ValueError:
        pass

    # Fallback: extract the first numeric-looking token from mixed copied text.
    m = re.search(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", s)
    if m:
        token = re.sub(r"(?<=\d)\.(?=[eE][+-]?\d+$)", "", m.group(0))
        try:
            return float(token)
        except ValueError:
            return default

    return default


# -----------------------------------------------------------------------------
# Text parser utilities
# -----------------------------------------------------------------------------

def read_text(path: str | Path) -> str:
    return Path(path).read_text(errors="ignore")


def clean_listing_lines(text: str) -> List[str]:
    """Remove MCNP echo-listing line numbers such as '804-       FS4006 ...'."""
    out: List[str] = []
    for raw in text.splitlines():
        m = re.match(r"^\s*\d+-\s{0,10}(.*)$", raw)
        out.append(m.group(1).rstrip() if m else raw.rstrip())
    return out


def parse_ints(text: str) -> List[int]:
    return [int(x) for x in re.findall(r"(?<![\d.])[-+]?\d+(?![\d.])", text)]


def parse_floats(text: str) -> List[float]:
    return [float(x) for x in re.findall(FLOAT_RE, text)]


def collect_cards(lines: List[str]) -> Dict[str, str]:
    """Collect named MCNP cards such as FC4006, F4006, FS4006, E5004, SD4006."""
    cards: Dict[str, str] = {}

    for i, line in enumerate(lines):
        s = line.strip()
        if not s or s.lower().startswith("c "):
            continue

        m = re.match(r"^([A-Za-z]+)(\d+)(?::[A-Za-z]+)?\s*(.*)$", s)
        if not m:
            continue

        prefix, num, rest = m.group(1).upper(), m.group(2), m.group(3)
        key = f"{prefix}{num}"
        chunks = [rest.split("$", 1)[0].strip()]
        j = i + 1

        while j < len(lines):
            t = lines[j].strip()
            if not t or t.lower().startswith("c "):
                break
            if re.match(r"^[A-Za-z]+\d+(?::[A-Za-z]+)?\b", t):
                break
            if re.match(r"^\d+\s+", t):
                break
            chunks.append(t.split("$", 1)[0].strip())
            j += 1

        cards[key] = " ".join(chunks).strip()

    return cards


def parse_surfaces(lines: List[str]) -> Dict[int, Dict[str, float | str]]:
    """Parse PZ surfaces as {surface_id: {'type': 'PZ', 'z': value}}."""
    surfaces: Dict[int, Dict[str, float | str]] = {}
    for line in lines:
        s = line.strip().split("$", 1)[0].strip()
        m = re.match(rf"^(\d+)\s+(PZ)\s+({FLOAT_RE})\b", s, re.I)
        if m:
            surfaces[int(m.group(1))] = {"type": "PZ", "z": float(m.group(3))}
    return surfaces


def parse_cell_cards(lines: List[str]) -> Dict[int, str]:
    """Parse echoed CELL CARDS section."""
    cells: Dict[int, str] = {}
    in_cells = False

    for i, line in enumerate(lines):
        s = line.strip()
        if "CELL CARDS" in s:
            in_cells = True
            continue
        if "SURFACE CARDS" in s:
            in_cells = False
            continue
        if not in_cells or not s or s.lower().startswith("c "):
            continue

        m = re.match(r"^(\d+)\s+(.+)$", s)
        if not m:
            continue

        cell_id = int(m.group(1))
        chunks = [s.split("$", 1)[0].strip()]
        j = i + 1
        while j < len(lines):
            t = lines[j].strip()
            if not t or t.lower().startswith("c "):
                break
            if re.match(r"^\d+\s+", t):
                break
            chunks.append(t.split("$", 1)[0].strip())
            j += 1
        cells[cell_id] = " ".join(chunks)

    return cells


def get_tally_result_block(text: str, tally_no: int) -> str:
    """Return final result block for tally_no, preferring the block containing 'nps ='."""
    patterns = [rf"1tally\s+{tally_no}\s+nps\b", rf"1tally\s+{tally_no}\b"]
    for pat in patterns:
        matches = list(re.finditer(pat, text))
        if matches:
            m = matches[-1]
            nxt = re.search(r"\n1tally\s+\d+", text[m.end():])
            end = m.end() + nxt.start() if nxt else len(text)
            return text[m.start():end]
    raise ValueError(f"Tally {tally_no} was not found in this outp file.")


def cell_z_range(cell_expr: str, surfaces: Dict[int, Dict[str, float | str]]) -> Tuple[float, float]:
    z_values: List[float] = []
    for signed_id in parse_ints(cell_expr):
        sid = abs(signed_id)
        if sid in surfaces and surfaces[sid].get("type") == "PZ":
            z_values.append(float(surfaces[sid]["z"]))
    if not z_values:
        raise ValueError(f"No PZ surfaces found in cell expression: {cell_expr}")
    return min(z_values), max(z_values)


def infer_cells_from_f_card(f_card: str) -> List[int]:
    """Best-effort cell list from F card text. Parenthesized unions are ignored."""
    # Remove inline parenthesized group to avoid duplicated union terms becoming first cell.
    s = re.sub(r"\([^)]*\)", " ", f_card)
    return [abs(i) for i in parse_ints(s)]


def get_tally_type_from_block(block: str) -> Optional[int]:
    m = re.search(r"tally type\s+(\d+)", block, re.I)
    return int(m.group(1)) if m else None



def get_fc_name_from_block(block: str) -> str:
    """Extract the printed FC/comment title from a tally result block if input cards are absent."""
    lines = block.splitlines()
    for line in lines[:12]:
        s = line.strip()
        if s.startswith("+"):
            return s[1:].strip()
    return ""


def block_has_energy_table(block: str) -> bool:
    """True when a tally result block contains one or more 'energy' tables."""
    return bool(re.search(r"^\s*energy\s*$", block, re.I | re.M))


def parse_cell_aliases_from_block(block: str) -> Dict[str, str]:
    """
    Parse MCNP output lines such as:
        cell  d is (101 102 103)

    Returns {"d": "(101 102 103)"}.
    """
    aliases: Dict[str, str] = {}
    for m in re.finditer(r"^\s*cell\s+(\S+)\s+is\s+(.+?)\s*$", block, re.I | re.M):
        aliases[m.group(1)] = m.group(2).strip()
    return aliases

def segment_z_bounds(
    cards: Dict[str, str],
    tally_no: int,
    cell_id: int,
    cells: Dict[int, str],
    surfaces: Dict[int, Dict[str, float | str]],
) -> List[Tuple[Optional[int], float]]:
    if cell_id not in cells:
        raise ValueError(f"Cell {cell_id} was not found in echoed CELL CARDS.")

    z_min, z_max = cell_z_range(cells[cell_id], surfaces)
    fs_ids = [abs(i) for i in parse_ints(cards.get(f"FS{tally_no}", ""))]

    cuts: List[Tuple[int, float]] = []
    for sid in fs_ids:
        if sid in surfaces and surfaces[sid].get("type") == "PZ":
            z = float(surfaces[sid]["z"])
            if z_min < z < z_max:
                cuts.append((sid, z))

    ordered: List[Tuple[int, float]] = []
    seen_z = set()
    for sid, z in sorted(cuts, key=lambda item: item[1]):
        key = round(z, 10)
        if key not in seen_z:
            ordered.append((sid, z))
            seen_z.add(key)

    return [(None, z_min)] + ordered + [(None, z_max)]


# -----------------------------------------------------------------------------
# Tally result extraction
# -----------------------------------------------------------------------------

@dataclass
class TallyInfo:
    file_path: Path
    tally_no: int
    fc_name: str
    f_card: str
    tally_type: Optional[int]
    has_fs: bool
    has_e: bool
    category: str


def list_tallies(path: str | Path) -> List[TallyInfo]:
    path = Path(path)
    text = read_text(path)
    lines = clean_listing_lines(text)
    cards = collect_cards(lines)

    tally_numbers = set()
    for key in cards:
        m = re.match(r"^(?:FC|F|FS|E|FM|SD|FMESH)(\d+)$", key)
        if m:
            tally_numbers.add(int(m.group(1)))

    # Also support result-only snippets such as test.txt that contain a printed
    # "1tally #### nps =" block but not the echoed input cards.
    for m in re.finditer(r"1tally\s+(\d+)\b", text):
        tally_numbers.add(int(m.group(1)))

    infos: List[TallyInfo] = []
    for no in sorted(tally_numbers):
        # FMESH tallies are intentionally excluded from the interactive tally list.
        # They have a different mesh-output structure and are usually better handled
        # by a dedicated mesh parser/visualizer rather than this cell/energy table UI.
        if f"FMESH{no}" in cards:
            continue

        block = ""
        tally_type = None
        try:
            block = get_tally_result_block(text, no)
            tally_type = get_tally_type_from_block(block)
        except Exception:
            pass

        fc = cards.get(f"FC{no}", "") or (get_fc_name_from_block(block) if block else "")
        f_card = cards.get(f"F{no}", "")
        has_fs = f"FS{no}" in cards
        has_e = f"E{no}" in cards or (block_has_energy_table(block) if block else False)

        if has_fs:
            category = "FS segment"
        elif has_e:
            category = "Energy bin"
        else:
            category = "Simple"

        infos.append(TallyInfo(path, no, fc, f_card, tally_type, has_fs, has_e, category))
    return infos


def parse_simple_tally(block: str) -> List[Dict[str, object]]:
    """Extract one-value tally. Works for TBR-like multiplier bins and plain cell values."""
    lines = block.splitlines()
    start_idx = 0
    for i, line in enumerate(lines):
        if any(key in line.lower() for key in ["multiplier bin", "cell "]):
            start_idx = i
            if "multiplier bin" in line.lower():
                break

    for line in lines[start_idx:]:
        m = re.match(rf"^\s*({FLOAT_RE})\s+({FLOAT_RE})\s*$", line)
        if m:
            return [{"tally_value": float(m.group(1)), "relative_error": float(m.group(2))}]
    raise ValueError("Could not find a simple tally value and relative error.")



def is_mcnp_value_error_line(line: str) -> Optional[Tuple[float, float]]:
    """Return (value, relative_error) only for MCNP result rows.

    This intentionally rejects FS segment continuation lines such as
        9139      -9140
    or
        9139       9140      -9141
    which otherwise look like two numeric fields.  A result row has exactly
    two numeric tokens, the second token is a non-negative relative error,
    and at least one token has MCNP floating-point notation (decimal point
    or exponent).
    """
    tokens = line.strip().split()
    if len(tokens) != 2:
        return None
    if not all(re.fullmatch(FLOAT_RE, t) for t in tokens):
        return None
    # Surface ID continuation lines are integer tokens. MCNP tally values and
    # relative errors are printed as decimal/exponent floats.
    if not any(("." in t or "e" in t.lower()) for t in tokens):
        return None
    try:
        value = float(tokens[0])
        relerr = float(tokens[1])
    except ValueError:
        return None
    if relerr < 0:
        return None
    return value, relerr

def parse_segment_tally(
    path: Path,
    tally_no: int,
    text: str,
    cards: Dict[str, str],
    cells: Dict[int, str],
    surfaces: Dict[int, Dict[str, float | str]],
) -> List[Dict[str, object]]:
    block = get_tally_result_block(text, tally_no)
    rows: List[Dict[str, object]] = []
    current_cell: Optional[int] = None
    collecting_segment = False
    segment_tokens: List[int] = []

    for raw in block.splitlines():
        line = raw.rstrip()
        m_cell = re.match(r"\s*cell\s+(\d+)\b", line)
        if m_cell:
            current_cell = int(m_cell.group(1))
            collecting_segment = False
            segment_tokens = []
            continue

        if "segment:" in line:
            collecting_segment = True
            segment_tokens = parse_ints(line.split("segment:", 1)[1])
            continue

        if collecting_segment:
            value_pair = is_mcnp_value_error_line(line)
            if value_pair is not None:
                value, relerr = value_pair
                rows.append(
                    {
                        "cell": current_cell,
                        "tally_value": value,
                        "relative_error": relerr,
                    }
                )
                collecting_segment = False
                segment_tokens = []
            else:
                # This is usually a wrapped continuation of the segment surface list,
                # e.g. "9139      -9140". Keep consuming it until the actual
                # two-column tally value/error row appears.
                segment_tokens.extend(parse_ints(line))

    if not rows:
        raise ValueError(f"No FS segment result rows were found for tally {tally_no}.")

    cell_id = int(rows[0]["cell"] or infer_cells_from_f_card(cards.get(f"F{tally_no}", ""))[0])
    bounds = segment_z_bounds(cards, tally_no, cell_id, cells, surfaces)

    out: List[Dict[str, object]] = []
    for i, row in enumerate(rows):
        z0 = bounds[i][1] if i < len(bounds) - 1 else None
        z1 = bounds[i + 1][1] if i + 1 < len(bounds) else None
        z_mid = (z0 + z1) / 2.0 if z0 is not None and z1 is not None else None
        out.append(
            {
                "file": path.name,
                "tally": tally_no,
                "fc_name": cards.get(f"FC{tally_no}", ""),
                "cell": cell_id,
                "z_start_cm": z0,
                "z_end_cm": z1,
                "z_mid_cm": z_mid,
                "tally_value": row["tally_value"],
                "relative_error": row["relative_error"],
            }
        )
    return out


def parse_energy_tally(
    path: Path,
    tally_no: int,
    text: str,
    cards: Dict[str, str],
) -> List[Dict[str, object]]:
    """
    Extract energy-bin rows from MCNP energy-table tally blocks.

    Supported result-bin headers
    ----------------------------
    This version supports both cell-based energy tallies such as F4 and
    surface-based energy tallies such as F1/F2.  MCNP prints these sections
    with headers such as:

        cell 401
        energy
        ...

        surface 200
        energy
        ...

    Energy-bin interpretation used in v12
    -------------------------------------
    MCNP energy-table rows have the form:

        E_upper   tally   relative_error

    The parser stores the printed tally as ``raw_tally`` and additionally
    computes two bin-normalized columns requested by the user:

        delta_E_MeV       = E_upper - E_lower
        log_delta_E       = ln(E_upper / E_lower)
        tally_per_delta_E = raw_tally / delta_E_MeV
        tally_per_lethargy = raw_tally / log_delta_E

    ``log_delta_E`` is dimensionless and corresponds to the lethargy bin width.
    When E_lower <= 0, log_delta_E and tally_per_lethargy are left blank.
    """
    block = get_tally_result_block(text, tally_no)
    e_bounds = parse_floats(cards.get(f"E{tally_no}", ""))
    aliases = parse_cell_aliases_from_block(block)
    fc_name = cards.get(f"FC{tally_no}", "") or get_fc_name_from_block(block)

    rows: List[Dict[str, object]] = []
    current_bin_type: Optional[str] = None
    current_bin: Optional[str] = None
    energy_row_index = 0
    previous_upper_by_bin: Dict[Tuple[str, str], float] = {}

    for line in block.splitlines():
        # Alias/definition lines such as "cell  d is (101 102 103)" are metadata,
        # not a new result table.
        m_alias = re.match(r"^\s*cell\s+(\S+)\s+is\s+(.+?)\s*$", line, re.I)
        if m_alias:
            aliases[m_alias.group(1)] = m_alias.group(2).strip()
            continue

        # Result table start. Supported labels include:
        #   cell 101
        #   cell d
        #   cell (101 102 103)
        #   surface 200
        m_bin = re.match(r"^\s*(cell|surface)\s+(.+?)\s*$", line, re.I)
        if m_bin:
            btype = m_bin.group(1).lower()
            label = m_bin.group(2).strip()

            # MCNP sometimes prints the union expression itself even after saying
            # "cell d is (101 102 103)". Convert it back to "d" when possible.
            if btype == "cell":
                for alias_key, alias_value in aliases.items():
                    if label == alias_value:
                        label = alias_key
                        break

            current_bin_type = btype
            current_bin = label
            energy_row_index = 0
            continue

        if re.match(r"^\s*energy\s*$", line, re.I):
            energy_row_index = 0
            continue

        # Total is not an energy bin, so delta_E/log_delta_E are intentionally blank.
        m_total = re.match(rf"^\s*total\s+({FLOAT_RE})\s+({FLOAT_RE})\s*$", line, re.I)
        if m_total and current_bin_type is not None and current_bin is not None:
            value = float(m_total.group(1))
            rel = float(m_total.group(2))
            row = {
                "file": path.name,
                "tally": tally_no,
                "fc_name": fc_name,
                "bin_type": current_bin_type,
                "bin": current_bin,
                "bin_definition": aliases.get(current_bin, "") if current_bin_type == "cell" else "",
                "energy_lower_MeV": None,
                "energy_upper_MeV": "total",
                "delta_E_MeV": None,
                "log_delta_E": None,
                "raw_tally": value,
                "tally_per_delta_E": None,
                "tally_per_lethargy": None,
                "relative_error": rel,
            }
            if current_bin_type == "cell":
                row["cell_bin"] = current_bin
                row["cell_bin_definition"] = aliases.get(current_bin, "")
            elif current_bin_type == "surface":
                row["surface_bin"] = current_bin
            rows.append(row)
            continue

        # MCNP energy-bin result lines generally look like:
        #   upper_energy   tally   relative_error
        m = re.match(rf"^\s*({FLOAT_RE})\s+({FLOAT_RE})\s+({FLOAT_RE})\s*$", line)
        if m and current_bin_type is not None and current_bin is not None:
            printed_upper = float(m.group(1))
            raw_tally = float(m.group(2))
            rel = float(m.group(3))

            if e_bounds:
                # E-card values are the printed upper boundaries.  The lower boundary
                # of the first bin is treated as 0.0 because MCNP output lists the
                # first E-card value as the first printed energy boundary.
                lower = 0.0 if energy_row_index == 0 else e_bounds[min(energy_row_index - 1, len(e_bounds) - 1)]
                upper = e_bounds[min(energy_row_index, len(e_bounds) - 1)]

                # If printed upper and parsed E-card upper differ meaningfully, use
                # the printed upper because it reflects the actual result row.
                if abs(upper - printed_upper) / max(abs(printed_upper), 1e-300) > 1e-6:
                    upper = printed_upper
            else:
                # For result-only files/snippets without E cards, infer lower bounds
                # from the previous printed upper energy within the same bin table.
                key = (current_bin_type, current_bin)
                lower = previous_upper_by_bin.get(key, 0.0)
                upper = printed_upper
                previous_upper_by_bin[key] = upper

            delta_E = None
            log_delta_E = None
            tally_per_delta_E = None
            tally_per_lethargy = None

            try:
                if lower is not None and upper is not None and float(upper) > float(lower):
                    delta_E = float(upper) - float(lower)
                    if delta_E > 0.0:
                        tally_per_delta_E = raw_tally / delta_E
                    if float(lower) > 0.0:
                        log_delta_E = math.log(float(upper) / float(lower))
                        if log_delta_E > 0.0:
                            tally_per_lethargy = raw_tally / log_delta_E
            except Exception:
                delta_E = None
                log_delta_E = None
                tally_per_delta_E = None
                tally_per_lethargy = None

            row = {
                "file": path.name,
                "tally": tally_no,
                "fc_name": fc_name,
                "bin_type": current_bin_type,
                "bin": current_bin,
                "bin_definition": aliases.get(current_bin, "") if current_bin_type == "cell" else "",
                "energy_lower_MeV": lower,
                "energy_upper_MeV": upper,
                "delta_E_MeV": delta_E,
                "log_delta_E": log_delta_E,
                "raw_tally": raw_tally,
                "tally_per_delta_E": tally_per_delta_E,
                "tally_per_lethargy": tally_per_lethargy,
                "relative_error": rel,
            }
            # Backward-compatible convenience columns.
            if current_bin_type == "cell":
                row["cell_bin"] = current_bin
                row["cell_bin_definition"] = aliases.get(current_bin, "")
            elif current_bin_type == "surface":
                row["surface_bin"] = current_bin
            rows.append(row)
            energy_row_index += 1

    if not rows:
        raise ValueError(
            f"No energy-bin rows were found for tally {tally_no}. "
            "This parser supports both 'cell ... energy' and 'surface ... energy' tables; "
            "please check whether the tally block uses another bin header."
        )
    return rows

def extract_tally(info: TallyInfo) -> List[Dict[str, object]]:
    text = read_text(info.file_path)
    lines = clean_listing_lines(text)
    cards = collect_cards(lines)
    cells = parse_cell_cards(lines)
    surfaces = parse_surfaces(lines)

    if info.has_fs:
        return parse_segment_tally(info.file_path, info.tally_no, text, cards, cells, surfaces)
    if info.has_e:
        return parse_energy_tally(info.file_path, info.tally_no, text, cards)

    block = get_tally_result_block(text, info.tally_no)
    simple_rows = parse_simple_tally(block)
    for row in simple_rows:
        row.update(
            {
                "file": info.file_path.name,
                "tally": info.tally_no,
                "fc_name": cards.get(f"FC{info.tally_no}", ""),
                "tally_type": get_tally_type_from_block(block),
            }
        )
    # Put metadata columns first.
    return [{k: row[k] for k in ["file", "tally", "fc_name", "tally_type", "tally_value", "relative_error"] if k in row} for row in simple_rows]




def _as_float_or_none(value):
    """Convert a value to float when possible; return None for blanks/NaN-like values."""
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        x = float(value)
    except Exception:
        return None
    if math.isnan(x) or math.isinf(x):
        return None
    return x


def apply_multiplier_offset(
    rows: List[Dict[str, object]],
    multiplier: float = 1.0,
    addend: float = 0.0,
) -> List[Dict[str, object]]:
    """Return rows with adjusted result columns while preserving original raw columns.

    The GUI uses this for unit conversion or simple post-processing, e.g.
    MeV/cm3 -> W/cm3. Original MCNP values are not overwritten.

    adjusted = raw tally * multiplier + addend

    Columns transformed when present:
    - tally_value: simple tally or F6/FS segment value
    - raw_tally: printed MCNP energy-bin tally value
    - tally_per_delta_E: raw_tally / delta_E_MeV
    - tally_per_lethargy: raw_tally / log_delta_E
    """
    result_columns = [
        "tally_value",
        "raw_tally",
        "tally_per_delta_E",
        "tally_per_lethargy",
    ]

    transformed: List[Dict[str, object]] = []
    for row in rows:
        new_row = dict(row)
        for col in result_columns:
            if col not in row:
                continue
            x = _as_float_or_none(row.get(col))
            new_row[f"{col}_adjusted"] = None if x is None else x * multiplier + addend

        # Keep the applied transform visible in both the table and exported files.
        new_row["applied_multiplier"] = multiplier
        new_row["applied_addend"] = addend
        transformed.append(new_row)
    return transformed



# -----------------------------------------------------------------------------
# Export helpers
# -----------------------------------------------------------------------------

def safe_excel_sheet_name(base: str, used: set[str]) -> str:
    """Return a unique Excel sheet name within Excel's 31-character limit."""
    # Excel disallows: : \/ ? * [ ]
    name = re.sub(r"[:\\/?*\[\]]", "_", base).strip()
    name = re.sub(r"\s+", " ", name)
    if not name:
        name = "Sheet"
    name = name[:31]

    candidate = name
    i = 1
    while candidate in used:
        suffix = f"_{i}"
        candidate = name[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate)
    return candidate


def tally_sheet_base_name(info: TallyInfo) -> str:
    fc = re.sub(r"\s+", "_", info.fc_name.strip()) if info.fc_name else info.category
    fc = re.sub(r"[^A-Za-z0-9가-힣_\-]+", "_", fc)
    if len(fc) > 14:
        fc = fc[:14]
    return f"{info.file_path.stem}_{info.tally_no}_{fc}"


# -----------------------------------------------------------------------------
# Qt GUI
# -----------------------------------------------------------------------------

class DataTableModelMixin:
    @staticmethod
    def rows_to_table(rows: List[Dict[str, object]]) -> Tuple[List[str], List[List[str]]]:
        if not rows:
            return [], []
        columns: List[str] = []
        for row in rows:
            for key in row.keys():
                if key not in columns:
                    columns.append(key)

        table: List[List[str]] = []
        for row in rows:
            table.append(["" if row.get(col) is None else str(row.get(col)) for col in columns])
        return columns, table


def main() -> None:
    try:
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import (
            QApplication,
            QAbstractItemView,
            QFileDialog,
            QHBoxLayout,
            QHeaderView,
            QLabel,
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QPushButton,
            QSplitter,
            QTableWidget,
            QTableWidgetItem,
            QTextEdit,
            QVBoxLayout,
            QWidget,
        )
    except ImportError as exc:
        print("PySide6 is not installed. Install it with: pip install PySide6 pandas openpyxl")
        raise exc

    class MainWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle(f"MCNP outp Tally Viewer v9 - {APP_VERSION}")
            self.resize(1300, 760)
            self.tallies: List[TallyInfo] = []
            self.visible_indices: List[int] = []
            self.current_raw_rows: List[Dict[str, object]] = []
            self.current_rows: List[Dict[str, object]] = []
            self.checked_tally_keys: set[tuple[str, int]] = set()
            self._populating_tally_table = False
            self._tally_sort_orders: dict[int, Qt.SortOrder] = {}

            root = QWidget()
            self.setCentralWidget(root)
            root_layout = QVBoxLayout(root)

            top = QHBoxLayout()
            self.open_btn = QPushButton("Open outp file(s)")
            self.export_btn = QPushButton("Export current table to CSV")
            self.check_visible_btn = QPushButton("Check visible tallies")
            self.batch_export_btn = QPushButton("Export checked tallies to XLSX")
            self.clear_checks_btn = QPushButton("Clear checks")
            self.filter_edit = QLineEdit()
            self.filter_edit.setPlaceholderText("Filter by tally number, FC name, file, category...")
            top.addWidget(self.open_btn)
            top.addWidget(self.export_btn)
            top.addWidget(self.check_visible_btn)
            top.addWidget(self.batch_export_btn)
            top.addWidget(self.clear_checks_btn)
            top.addWidget(self.filter_edit, 1)
            root_layout.addLayout(top)

            splitter = QSplitter(Qt.Horizontal)
            root_layout.addWidget(splitter, 1)

            left_widget = QWidget()
            left_layout = QVBoxLayout(left_widget)
            left_layout.addWidget(QLabel("Tally list"))
            self.tally_table = QTableWidget(0, 7)
            self.tally_table.setHorizontalHeaderLabels(["Export", "File", "Tally", "Type", "Category", "FC name", "F card"])
            self.tally_table.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.tally_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tally_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self.tally_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            self.tally_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
            self.tally_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
            self.tally_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
            self.tally_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
            self.tally_table.horizontalHeader().setSectionsClickable(True)
            self.tally_table.horizontalHeader().sectionClicked.connect(self.sort_tally_table_by_column)
            left_layout.addWidget(self.tally_table)
            splitter.addWidget(left_widget)

            right_widget = QWidget()
            right_layout = QVBoxLayout(right_widget)
            self.info_label = QLabel("Open one or more MCNP outp files, then select a tally.")
            self.info_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

            transform_layout = QHBoxLayout()
            transform_layout.addWidget(QLabel("Multiplier:"))
            self.multiplier_edit = QLineEdit("1")
            self.multiplier_edit.setPlaceholderText("e.g., 1.0, 2.5E-3")
            self.multiplier_edit.setMaximumWidth(130)
            transform_layout.addWidget(self.multiplier_edit)
            transform_layout.addWidget(QLabel("Add/Offset:"))
            self.addend_edit = QLineEdit("0")
            self.addend_edit.setPlaceholderText("e.g., 0")
            self.addend_edit.setMaximumWidth(130)
            transform_layout.addWidget(self.addend_edit)
            transform_layout.addWidget(QLabel("Adjusted = Raw tally × Multiplier + Add/Offset"))
            transform_layout.addSpacing(18)
            transform_layout.addWidget(QLabel("Preview rows:"))
            self.preview_limit_edit = QLineEdit(str(PREVIEW_ROW_LIMIT_DEFAULT))
            self.preview_limit_edit.setPlaceholderText("e.g., 500, all")
            self.preview_limit_edit.setMaximumWidth(95)
            transform_layout.addWidget(self.preview_limit_edit)
            transform_layout.addWidget(QLabel("(export saves all rows)"))
            transform_layout.addStretch(1)

            self.preview_label = QLabel("")
            self.preview_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

            self.result_table = QTableWidget(0, 0)
            self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
            self.log_box = QTextEdit()
            self.log_box.setReadOnly(True)
            self.log_box.setMaximumHeight(110)
            right_layout.addWidget(self.info_label)
            right_layout.addLayout(transform_layout)
            right_layout.addWidget(self.preview_label)
            right_layout.addWidget(self.result_table, 1)
            right_layout.addWidget(QLabel("Log"))
            right_layout.addWidget(self.log_box)
            splitter.addWidget(right_widget)
            splitter.setSizes([470, 830])

            self.open_btn.clicked.connect(self.open_files)
            self.export_btn.clicked.connect(self.export_current_table)
            self.check_visible_btn.clicked.connect(self.check_visible_tallies)
            self.batch_export_btn.clicked.connect(self.export_checked_tallies_to_xlsx)
            self.clear_checks_btn.clicked.connect(self.clear_checks)
            self.filter_edit.textChanged.connect(self.apply_filter)
            self.tally_table.itemSelectionChanged.connect(self.on_tally_selected)
            self.tally_table.itemChanged.connect(self.on_tally_item_changed)
            self.multiplier_edit.editingFinished.connect(self.refresh_current_transform)
            self.addend_edit.editingFinished.connect(self.refresh_current_transform)
            self.preview_limit_edit.editingFinished.connect(self.refresh_preview_only)

        def log(self, message: str) -> None:
            self.log_box.append(message)

        def open_files(self) -> None:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Open MCNP outp file(s)",
                "",
                "Text files (*.txt *.outp *.out *.o *.*)",
            )
            if not files:
                return

            loaded = 0
            for file in files:
                try:
                    infos = list_tallies(file)
                    self.tallies.extend(infos)
                    loaded += len(infos)
                    self.log(f"Loaded {len(infos)} tallies from {Path(file).name}")
                except Exception as exc:
                    QMessageBox.warning(self, "Load error", f"Failed to read {file}\n\n{exc}")
                    self.log(f"ERROR loading {file}: {exc}")

            self.apply_filter()
            self.info_label.setText(f"Loaded {loaded} tallies. Select a tally from the left table.")

        def tally_key(self, info: TallyInfo) -> tuple[str, int]:
            return (str(info.file_path.resolve()), int(info.tally_no))

        def apply_filter(self) -> None:
            needle = self.filter_edit.text().strip().lower()
            self.visible_indices = []
            for idx, info in enumerate(self.tallies):
                hay = " ".join(
                    [
                        info.file_path.name,
                        str(info.tally_no),
                        str(info.tally_type or ""),
                        info.category,
                        info.fc_name,
                        info.f_card,
                    ]
                ).lower()
                if not needle or needle in hay:
                    self.visible_indices.append(idx)

            self._populating_tally_table = True
            self.tally_table.setSortingEnabled(False)
            self.tally_table.setRowCount(len(self.visible_indices))
            for row, idx in enumerate(self.visible_indices):
                info = self.tallies[idx]

                check_item = QTableWidgetItem("")
                check_item.setFlags(check_item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                check_item.setCheckState(Qt.Checked if self.tally_key(info) in self.checked_tally_keys else Qt.Unchecked)
                check_item.setData(Qt.UserRole, idx)
                self.tally_table.setItem(row, 0, check_item)

                values = [
                    info.file_path.name,
                    str(info.tally_no),
                    "" if info.tally_type is None else str(info.tally_type),
                    info.category,
                    info.fc_name,
                    info.f_card,
                ]
                for offset, value in enumerate(values, start=1):
                    item = QTableWidgetItem(value)
                    item.setData(Qt.UserRole, idx)
                    # Numeric-aware sorting for Tally and Type columns.
                    if offset == 2:
                        item.setData(Qt.EditRole, int(info.tally_no))
                    elif offset == 3 and info.tally_type is not None:
                        item.setData(Qt.EditRole, int(info.tally_type))
                    self.tally_table.setItem(row, offset, item)
            self._populating_tally_table = False
            self.tally_table.resizeColumnsToContents()

        def sort_tally_table_by_column(self, column: int) -> None:
            """Sort the left tally list by selected header columns only.

            Sortable columns:
              1 File, 2 Tally, 3 Type, 4 Category, 5 FC name
            Export checkbox and F card columns are intentionally not sorted by header click.
            """
            sortable_columns = {1, 2, 3, 4, 5}
            if column not in sortable_columns:
                return

            previous = self._tally_sort_orders.get(column, Qt.DescendingOrder)
            order = Qt.AscendingOrder if previous == Qt.DescendingOrder else Qt.DescendingOrder
            self._tally_sort_orders[column] = order

            self.tally_table.sortItems(column, order)

            direction = "ascending" if order == Qt.AscendingOrder else "descending"
            header_name = self.tally_table.horizontalHeaderItem(column).text()
            self.log(f"Sorted tally list by {header_name} ({direction}).")


        def on_tally_item_changed(self, item: QTableWidgetItem) -> None:
            if self._populating_tally_table or item.column() != 0:
                return
            idx = item.data(Qt.UserRole)
            if idx is None:
                return
            info = self.tallies[int(idx)]
            key = self.tally_key(info)
            if item.checkState() == Qt.Checked:
                self.checked_tally_keys.add(key)
            else:
                self.checked_tally_keys.discard(key)

        def clear_checks(self) -> None:
            self.checked_tally_keys.clear()
            self.apply_filter()
            self.log("Cleared all checked tallies.")

        def check_visible_tallies(self) -> None:
            """Check only the tallies currently visible in the left list after filtering."""
            if not self.visible_indices:
                QMessageBox.information(self, "No visible tallies", "There are no visible tallies to check.")
                return

            for idx in self.visible_indices:
                info = self.tallies[idx]
                self.checked_tally_keys.add(self.tally_key(info))

            # Rebuild the table so check states are refreshed. Hidden rows remain untouched.
            n_checked_now = len(self.visible_indices)
            self.apply_filter()
            self.log(f"Checked {n_checked_now} currently visible tally/tallies.")

        def checked_infos(self) -> List[TallyInfo]:
            checked = []
            for info in self.tallies:
                if self.tally_key(info) in self.checked_tally_keys:
                    checked.append(info)
            return checked

        def transform_parameters(self) -> tuple[float, float]:
            """Read multiplier and addend from the UI.

            Accepts Excel-pasted formats such as 1,234.56 and 1.123.E+2.
            """
            multiplier = parse_user_number(self.multiplier_edit.text(), 1.0)
            addend = parse_user_number(self.addend_edit.text(), 0.0)

            # Normalize the displayed text so the user can see what was interpreted.
            self.multiplier_edit.setText(f"{multiplier:.12g}")
            self.addend_edit.setText(f"{addend:.12g}")

            return multiplier, addend

        def transformed_rows(self, raw_rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
            multiplier, addend = self.transform_parameters()
            return apply_multiplier_offset(raw_rows, multiplier, addend)

        def refresh_current_transform(self) -> None:
            """Re-apply multiplier/addend to the currently displayed tally."""
            if not self.current_raw_rows:
                return
            self.current_rows = self.transformed_rows(self.current_raw_rows)
            self.show_rows(self.current_rows)
            self.log(
                f"Updated displayed transform: multiplier={self.multiplier_edit.text().strip()}, "
                f"add/offset={self.addend_edit.text().strip()}"
            )

        def on_tally_selected(self) -> None:
            selected = self.tally_table.selectedItems()
            if not selected:
                return
            row = selected[0].row()
            item = self.tally_table.item(row, 0)
            idx = item.data(Qt.UserRole)
            info = self.tallies[idx]

            try:
                raw_rows = extract_tally(info)
                self.current_raw_rows = raw_rows
                rows = self.transformed_rows(raw_rows)
                self.current_rows = rows
                self.show_rows(rows)
                self.info_label.setText(
                    f"{info.file_path.name} | Tally {info.tally_no} | {info.category} | {info.fc_name}"
                )
                self.log(f"Extracted {len(rows)} row(s) from tally {info.tally_no} ({info.category}).")
            except Exception as exc:
                self.current_raw_rows = []
                self.current_rows = []
                self.result_table.setRowCount(0)
                self.result_table.setColumnCount(0)
                self.info_label.setText(f"Failed to extract tally {info.tally_no}: {exc}")
                self.log(f"ERROR extracting tally {info.tally_no}: {exc}")

        def preview_row_limit(self) -> Optional[int]:
            """Return the maximum number of rows to render in the Qt table.

            Rendering thousands of QTableWidgetItem objects can make the GUI feel frozen,
            especially for energy-bin tallies with multiple cell bins. Export functions still
            use self.current_rows, so they save the full table regardless of this preview limit.
            """
            text = self.preview_limit_edit.text().strip().lower()
            if text in {"", "all", "none", "0"}:
                return None
            try:
                value = int(float(text))
            except Exception:
                QMessageBox.warning(
                    self,
                    "Invalid preview rows",
                    f"Preview rows must be a positive integer or 'all'. It will be reset to {PREVIEW_ROW_LIMIT_DEFAULT}.",
                )
                self.preview_limit_edit.setText(str(PREVIEW_ROW_LIMIT_DEFAULT))
                return PREVIEW_ROW_LIMIT_DEFAULT
            if value <= 0:
                return None
            return value

        def refresh_preview_only(self) -> None:
            """Re-render only the current table preview without re-extracting the tally."""
            if self.current_rows:
                self.show_rows(self.current_rows)

        def show_rows(self, rows: List[Dict[str, object]]) -> None:
            columns, table = DataTableModelMixin.rows_to_table(rows)
            total_rows = len(table)
            limit = self.preview_row_limit()
            if limit is not None and total_rows > limit:
                preview_table = table[:limit]
                self.preview_label.setText(
                    f"Preview: showing first {limit:,} of {total_rows:,} rows. CSV/XLSX export saves all {total_rows:,} rows."
                )
            else:
                preview_table = table
                if total_rows:
                    self.preview_label.setText(f"Preview: showing all {total_rows:,} rows.")
                else:
                    self.preview_label.setText("Preview: no rows.")

            self.result_table.setUpdatesEnabled(False)
            try:
                self.result_table.clear()
                self.result_table.setColumnCount(len(columns))
                self.result_table.setHorizontalHeaderLabels(columns)
                self.result_table.setRowCount(len(preview_table))

                for r, row_values in enumerate(preview_table):
                    for c, value in enumerate(row_values):
                        self.result_table.setItem(r, c, QTableWidgetItem(value))
                # Resize based on the preview only, not the full extracted table.
                self.result_table.resizeColumnsToContents()
            finally:
                self.result_table.setUpdatesEnabled(True)

        def export_checked_tallies_to_xlsx(self) -> None:
            infos = self.checked_infos()
            if not infos:
                QMessageBox.information(self, "No checked tallies", "Check one or more tallies in the left table first.")
                return

            path, _ = QFileDialog.getSaveFileName(
                self,
                "Save checked tallies as Excel workbook",
                "mcnp_checked_tallies.xlsx",
                "Excel workbook (*.xlsx)",
            )
            if not path:
                return
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"

            try:
                import pandas as pd
            except ImportError as exc:
                QMessageBox.warning(
                    self,
                    "Missing package",
                    "pandas/openpyxl are required for XLSX export. Install with:\n\npip install pandas openpyxl",
                )
                raise exc

            used_sheet_names: set[str] = set()
            exported = 0
            failed: List[str] = []
            extracted_tables: List[Tuple[TallyInfo, str, List[Dict[str, object]]]] = []
            stacked_rows: List[Dict[str, object]] = []

            def beautify_sheet(ws) -> None:
                """Freeze the header, add filters, and set reasonable column widths."""
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for col_cells in ws.columns:
                    header = str(col_cells[0].value) if col_cells[0].value is not None else ""
                    max_len = max([len(str(c.value)) if c.value is not None else 0 for c in col_cells[:200]] + [len(header)])
                    ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 10), 34)

            # Extract first, so the workbook can place the stacked summary sheet first.
            for info in infos:
                try:
                    raw_rows = extract_tally(info)
                    rows = self.transformed_rows(raw_rows)
                    if not rows:
                        failed.append(f"{info.file_path.name} tally {info.tally_no}: no rows")
                        continue

                    sheet = safe_excel_sheet_name(tally_sheet_base_name(info), used_sheet_names)
                    extracted_tables.append((info, sheet, rows))

                    # Add a few batch-export columns so the stacked sheet remains readable
                    # even when different tally types have different column structures.
                    for row in rows:
                        stacked_row = {
                            "export_sheet": sheet,
                            "category": info.category,
                            "tally_type": info.tally_type,
                        }
                        stacked_row.update(row)
                        stacked_rows.append(stacked_row)
                except Exception as exc:
                    failed.append(f"{info.file_path.name} tally {info.tally_no}: {exc}")

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                # First sheet: all checked tallies stacked into one table.
                if stacked_rows:
                    summary_sheet = safe_excel_sheet_name("ALL_CHECKED_TALLIES", used_sheet_names)
                    all_df = pd.DataFrame(stacked_rows)
                    all_df.to_excel(writer, sheet_name=summary_sheet, index=False)
                    beautify_sheet(writer.book[summary_sheet])

                # Following sheets: one table per checked tally, as before.
                for info, sheet, rows in extracted_tables:
                    df = pd.DataFrame(rows)
                    df.to_excel(writer, sheet_name=sheet, index=False)
                    beautify_sheet(writer.book[sheet])
                    exported += 1

            if failed:
                self.log("Some checked tallies failed during XLSX export:")
                for msg in failed:
                    self.log("  " + msg)

            self.log(f"Saved XLSX with 1 stacked sheet + {exported} individual tally sheet(s): {path}")
            QMessageBox.information(
                self,
                "Saved",
                f"Excel workbook saved:\n{path}\n\nStacked summary sheet: 1\nIndividual tally sheets: {exported}\nFailed: {len(failed)}",
            )

        def export_current_table(self) -> None:
            if not self.current_rows:
                QMessageBox.information(self, "No data", "There is no currently displayed table to export.")
                return

            path, _ = QFileDialog.getSaveFileName(
                self,
                "Save CSV",
                "mcnp_tally_export.csv",
                "CSV files (*.csv)",
            )
            if not path:
                return

            columns, table = DataTableModelMixin.rows_to_table(self.current_rows)
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(columns)
                writer.writerows(table)
            self.log(f"Saved CSV: {path}")
            QMessageBox.information(self, "Saved", f"CSV saved:\n{path}")

    app = QApplication(sys.argv)
    # Make the viewer easier to read on high-DPI monitors.
    # You can change UI_FONT_POINT_SIZE near the top of this file if needed.
    app.setStyleSheet(f"""
        QWidget {{ font-size: {UI_FONT_POINT_SIZE}pt; }}
        QTableWidget {{ font-size: {UI_FONT_POINT_SIZE}pt; }}
        QHeaderView::section {{ font-size: {UI_FONT_POINT_SIZE}pt; font-weight: 600; padding: 4px; }}
        QPushButton {{ padding: 5px 9px; }}
        QLineEdit {{ padding: 4px; }}
    """)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
