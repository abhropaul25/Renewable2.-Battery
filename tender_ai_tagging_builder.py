#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tender_ai_tagging_builder.py

Builds an AI-tagging Excel from power-sector tender PDFs using a Lakhwar-style template.
- Clones the sheet structure from your template
- Extracts parameters via configurable regex rules (YAML)
- Populates AI_Tagging_Master (+ BID_INFO) with clause-level entries
- Updates AmendmentTracker with corrigenda/addenda
- Maintains versioning metadata in TenderMeta

Author: (You)
"""

import argparse
import re
import sys
import json
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Any, Optional
from pathlib import Path
import datetime as dt

import pandas as pd
from PyPDF2 import PdfReader

# ------------------------------
# Utilities
# ------------------------------

def log(msg: str):
    print(f"[tender-ai] {msg}", flush=True)

def read_pdf_text_with_pages(pdf_path: Path) -> List[Tuple[int, str]]:
    """Return list of (page_number, text) from a PDF. page_number is 1-indexed."""
    pages = []
    try:
        reader = PdfReader(str(pdf_path))
        for i, page in enumerate(reader.pages):
            try:
                t = page.extract_text() or ""
            except Exception:
                t = ""
            pages.append((i+1, t))
    except Exception as e:
        log(f"ERROR reading {pdf_path}: {e}")
    return pages

def clean_text(s: Any) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def find_first(pattern: str, pages: List[Tuple[int, str]], flags=re.IGNORECASE) -> Tuple[Optional[int], Optional[re.Match]]:
    rx = re.compile(pattern, flags)
    for pn, text in pages:
        m = rx.search(text)
        if m:
            return pn, m
    return None, None

def find_all(pattern: str, pages: List[Tuple[int, str]], flags=re.IGNORECASE) -> List[Tuple[int, re.Match]]:
    rx = re.compile(pattern, flags)
    hits = []
    for pn, text in pages:
        for m in rx.finditer(text):
            hits.append((pn, m))
    return hits

def to_int(s: str) -> Optional[int]:
    try:
        return int(re.sub(r"[^\d]", "", str(s)))
    except Exception:
        return None

def to_float(s: str) -> Optional[float]:
    try:
        return float(re.sub(r"[^0-9.\-]", "", str(s)))
    except Exception:
        return None

# ------------------------------
# YAML rule loading
# ------------------------------
try:
    import yaml
except Exception:
    yaml = None

def load_rules(yaml_path: Optional[Path]) -> Dict[str, Any]:
    if yaml_path is None:
        return {"extractors": [], "bid_info_map": {}, "defaults": {}}
    if yaml is None:
        raise RuntimeError("PyYAML not installed. Install with: pip install pyyaml")
    with open(yaml_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    # normalize shapes
    data.setdefault("extractors", [])
    data.setdefault("bid_info_map", {})
    data.setdefault("defaults", {})
    return data

# ------------------------------
# Data classes
# ------------------------------

@dataclass
class TagRow:
    Section: str
    ClauseRef: str
    Parameter: str
    Value: Any
    Unit: str = ""
    Notes: str = ""
    SourcePage: Any = ""

    def to_dict(self):
        return {
            "Section": self.Section,
            "Clause/Ref": self.ClauseRef,
            "Parameter": self.Parameter,
            "Value": self.Value,
            "Unit": self.Unit,
            "Notes": self.Notes,
            "SourcePage": self.SourcePage,
        }

# ------------------------------
# Core extraction logic
# ------------------------------

def run_extractors(pages: List[Tuple[int, str]], extractors: List[Dict[str, Any]]) -> List[TagRow]:
    """
    Each extractor item in YAML supports keys:
      - section: str
      - clause: str
      - parameter: str
      - unit: str (optional)
      - notes: str (optional)
      - mode: "first" | "all" (default first)
      - pattern: regex string, with named or positional groups
      - value_expr: Python format (e.g. "{0}/{1} MWh" or "{cap_mw}") using groups by index or name
      - flags: list of strings among ["IGNORECASE","MULTILINE","DOTALL"] (optional)
    """
    results: List[TagRow] = []
    for ext in extractors:
        section = ext.get("section","General")
        clause  = ext.get("clause","")
        parameter = ext.get("parameter","")
        unit = ext.get("unit","")
        notes = ext.get("notes","")
        mode = ext.get("mode","first")
        pattern = ext.get("pattern","")
        value_expr = ext.get("value_expr","{0}")
        flags_list = ext.get("flags",["IGNORECASE"])

        # compile flags
        fl = 0
        for f in flags_list:
            fl |= getattr(re, f, 0)

        # search
        if mode == "all":
            hits = find_all(pattern, pages, flags=fl)
            for pn, m in hits:
                value = render_value(value_expr, m)
                results.append(TagRow(section, clause, parameter, value, unit, notes, pn))
        else:
            pn, m = find_first(pattern, pages, flags=fl)
            if m:
                value = render_value(value_expr, m)
                results.append(TagRow(section, clause, parameter, value, unit, notes, pn))
    return results

def render_value(value_expr: str, m: re.Match) -> str:
    # Build named+positional dict
    fmt_dict = {}
    # positional: {0}, {1}, ...
    for i, g in enumerate(m.groups() or []):
        fmt_dict[str(i)] = clean_text(g)
    # named groups
    if hasattr(m, "groupdict"):
        for k, v in m.groupdict().items():
            fmt_dict[k] = clean_text(v)
    try:
        return value_expr.format(**fmt_dict, **{str(k):v for k,v in fmt_dict.items()})
    except Exception:
        # fallback: first group or full match
        return clean_text(m.group(1) if m.groups() else m.group(0))

# ------------------------------
# Workbook operations
# ------------------------------

def clone_template(template_xlsx: Path) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(template_xlsx)
    clones = {sn: pd.read_excel(xls, sheet_name=sn, dtype=object).iloc[0:0] for sn in xls.sheet_names}
    return clones

def read_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name, dtype=object)
    except Exception:
        return pd.DataFrame()

def upsert_bid_info(bid_info_df: pd.DataFrame, mapping: Dict[str, Any]) -> pd.DataFrame:
    if bid_info_df.empty:
        bid_info_df = pd.DataFrame(columns=["Field","Value"])
    for k, v in mapping.items():
        if (bid_info_df["Field"] == k).any():
            bid_info_df.loc[bid_info_df["Field"]==k, "Value"] = v
        else:
            bid_info_df.loc[len(bid_info_df)] = {"Field": k, "Value": v}
    return bid_info_df

def make_version_string(base: Optional[str]=None) -> str:
    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return f"{base or 'AI_Tagging'} — built {ts}"

def append_amendments(amendment_files: List[Path]) -> pd.DataFrame:
    rows = []
    for f in amendment_files:
        pages = read_pdf_text_with_pages(f)
        # try to detect type & date
        label = "Corrigendum/Addendum"
        # naive date grep
        pn, m = find_first(r"(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})", pages)
        date_found = m.group(1) if m else ""
        pn2, m2 = find_first(r"(Corrigendum|Addendum|Amendment)", pages)
        if m2:
            label = clean_text(m2.group(1))
        rows.append({
            "AmendmentType": label,
            "FileName": f.name,
            "Date": date_found,
            "Notes": "",
            "Pages": len(pages)
        })
    return pd.DataFrame(rows, columns=["AmendmentType","FileName","Date","Notes","Pages"])

# ------------------------------
# Main pipeline
# ------------------------------

def main():
    ap = argparse.ArgumentParser(description="Build AI-tagging Excel from tender PDFs")
    ap.add_argument("--template", required=True, help="Path to Lakhwar-style Excel template")
    ap.add_argument("--pdf", action="append", required=True, help="Path(s) to main tender PDFs (RfS/ITB etc.). Use multiple --pdf args to include more")
    ap.add_argument("--amendment", action="append", default=[], help="Path(s) to corrigendum/addendum PDFs")
    ap.add_argument("--rules", default=None, help="YAML rules file for regex extractors")
    ap.add_argument("--out", required=True, help="Output Excel path")
    args = ap.parse_args()

    template_xlsx = Path(args.template)
    out_xlsx = Path(args.out)
    main_pdfs = [Path(p) for p in args.pdf]
    amend_pdfs = [Path(p) for p in args.amendment]

    # Load rules
    rules = load_rules(Path(args.rules)) if args.rules else {"extractors": [], "bid_info_map": {}, "defaults": {}}

    # Clone template structure (preserves your schema/sheets)
    log("Cloning template structure...")
    clones = clone_template(template_xlsx)

    # Collect all pages from main PDFs
    log("Reading main PDFs...")
    all_pages: List[Tuple[int,str,str]] = []  # (page, text, source_file)
    for pdf in main_pdfs:
        pages = read_pdf_text_with_pages(pdf)
        for pn, tx in pages:
            all_pages.append((pn, tx, pdf.name))

    # Build a (page,text) list for extractors (ignoring filename)
    pages_simple = [(pn, tx) for pn, tx, _ in all_pages]

    # Run extractors -> TagRows
    log("Running extractors...")
    tag_rows = run_extractors(pages_simple, rules.get("extractors", []))

    # Build AI_Tagging_Master
    master_cols = ["Section","Clause/Ref","Parameter","Value","Unit","Notes","SourcePage"]
    df_master = pd.DataFrame([tr.to_dict() for tr in tag_rows], columns=master_cols)

    # BID_INFO from rule mapping (key -> regex pattern to find first match)
    bid_info_map: Dict[str, str] = rules.get("bid_info_map", {})
    bid_info_values: Dict[str, Any] = {}
    for field, pattern in bid_info_map.items():
        pn, m = find_first(pattern, pages_simple, flags=re.IGNORECASE | re.DOTALL | re.MULTILINE)
        bid_info_values[field] = clean_text(m.group(1)) if m else ""

    # Apply defaults if empty
    for k, v in rules.get("defaults", {}).items():
        bid_info_values.setdefault(k, v)

    df_bidinfo = upsert_bid_info(pd.DataFrame(columns=["Field","Value"]), bid_info_values)

    # Versioning & TenderMeta
    df_meta = pd.DataFrame([
        {"Key":"Build", "Value": make_version_string()},
        {"Key":"SourceCount", "Value": str(len(main_pdfs))},
        {"Key":"AmendmentsCount", "Value": str(len(amend_pdfs))},
    ])

    # AmendmentTracker
    df_amend = append_amendments(amend_pdfs)

    # Write output: clone all template sheets (empty), then overwrite key sheets if exist
    log(f"Writing output -> {out_xlsx}")
    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as writer:
        for sn, df in clones.items():
            safe = sn[:31]
            df.to_excel(writer, index=False, sheet_name=safe)
        # Write/merge standard sheets
        df_master.to_excel(writer, index=False, sheet_name="AI_Tagging_Master")
        df_bidinfo.to_excel(writer, index=False, sheet_name="BID_INFO")
        df_amend.to_excel(writer, index=False, sheet_name="AmendmentTracker")
        df_meta.to_excel(writer, index=False, sheet_name="TenderMeta")

    log("Done.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
