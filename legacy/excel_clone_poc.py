#!/usr/bin/env python3
"""
POC: ZIP-level Saving_Calculations sheet cloner.

Clones the Saving_Calculations sheet from an ABB appraisal template workbook,
producing a second sheet (Saving_Calculations_2) that is a structural replica:
  - renamed Excel table (Saving_Table2410 → Saving_Table2410_2)
  - all ~13,883 structured references updated to the new table name
  - drawing / chart links stripped  (charts remain on the original sheet)
  - calcChain.xml removed           (Excel rebuilds on first open)

Design ref: ZIP-level cloning design doc (2026-05-29).
POC scope only — no data filling, no UI, no production hardening.

Usage:
    python legacy/excel_clone_poc.py
    python legacy/excel_clone_poc.py <template.xlsx> <output.xlsx>
"""

import io
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

# ---------------------------------------------------------------------------
# OOXML constants
# ---------------------------------------------------------------------------

REL_WORKSHEET = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
REL_TABLE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
)
REL_DRAWING = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
REL_VMLDRAWING = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
)

CT_WORKSHEET = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
CT_TABLE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
)

CLONE_SUFFIX = "_2"

DEFAULT_TEMPLATE = (
    "legacy/excel_templates/1_0_region_site_name_saving_calculations.xlsx"
)
DEFAULT_OUTPUT = "legacy/excel_clone_poc_output.xlsx"

# ---------------------------------------------------------------------------
# Step 1 — load all ZIP parts into memory
# ---------------------------------------------------------------------------

def _load_parts(path: str) -> dict:
    parts = {}
    with zipfile.ZipFile(path, "r") as zf:
        for name in zf.namelist():
            parts[name] = zf.read(name)
    return parts


# ---------------------------------------------------------------------------
# Step 2 — discovery helpers
# ---------------------------------------------------------------------------

def _find_sheet_part(parts: dict) -> tuple:
    """
    Locate the Saving_Calculations sheet by inspecting workbook.xml and its
    relationship file.  Works for both 'Saving_Calculations' (v1/v3) and
    'Saving Calculations' (v2).

    Returns: (part_path, rels_path, display_name)
    """
    wb_text = parts["xl/workbook.xml"].decode()

    # Match <sheet name="Saving..." r:id="rIdN"/>
    m = re.search(
        r'<sheet\b[^>]+name="([^"]*[Ss]aving[^"]*)"[^>]+r:id="([^"]+)"',
        wb_text,
    )
    if not m:
        raise ValueError("No Saving… sheet found in workbook.xml")

    display_name = m.group(1)
    rel_id = m.group(2)

    # Resolve r:id → part path via workbook.xml.rels
    wb_rels = parts["xl/_rels/workbook.xml.rels"].decode()
    tm = re.search(
        r'<Relationship\b[^>]+Id="' + re.escape(rel_id) + r'"[^>]+Target="([^"]+)"',
        wb_rels,
    )
    if not tm:
        raise ValueError(f"Relationship {rel_id} not found in workbook.xml.rels")

    # target is relative from xl/: e.g. "worksheets/sheet3.xml"
    target = tm.group(1)
    part_path = "xl/" + target

    # _rels path: xl/worksheets/_rels/sheet3.xml.rels
    base = os.path.basename(target)          # "sheet3.xml"
    rels_path = f"xl/worksheets/_rels/{base}.rels"

    return part_path, rels_path, display_name


def _find_table_parts(parts: dict, sheet_rels_path: str, sheet_path: str) -> tuple:
    """
    Find the two Excel tables linked from the sheet's _rels file.
    Classifies them: main table = the one whose name[…] appears in sheet
    formula content; summary = the other.

    Returns:
        ((main_zip_path, main_name), (summ_zip_path, summ_name))
    """
    rels_text = parts[sheet_rels_path].decode()

    # Extract all Relationship elements, filter by table type
    raw_targets = []
    for attr_block in re.findall(r"<Relationship\b([^>]+)/>", rels_text):
        if REL_TABLE in attr_block:
            tm = re.search(r'Target="([^"]+)"', attr_block)
            if tm:
                raw_targets.append(tm.group(1))

    if not raw_targets:
        raise ValueError(f"No table relationships in {sheet_rels_path}")

    # Resolve relative targets → zip part paths
    # Targets are relative to xl/worksheets/:  "../tables/table1.xml"
    resolved = []
    for t in raw_targets:
        zip_path = os.path.normpath("xl/worksheets/" + t).replace("\\", "/")
        tname = re.search(rb'\bname="([^"]+)"', parts[zip_path]).group(1).decode()
        resolved.append((zip_path, tname))

    sheet_bytes = parts[sheet_path]
    main = None
    summ = None
    for zip_path, tname in resolved:
        if (tname + "[").encode() in sheet_bytes:
            main = (zip_path, tname)
        else:
            summ = (zip_path, tname)

    # Fallback: use declaration order if neither name appears with [
    if main is None:
        main, summ = resolved[0], resolved[1]

    return main, summ


def _next_sheet_part(parts: dict) -> str:
    """Return the next available xl/worksheets/sheetN.xml path."""
    nums = []
    for name in parts:
        m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", name)
        if m:
            nums.append(int(m.group(1)))
    return f"xl/worksheets/sheet{max(nums) + 1}.xml"


def _next_table_parts(parts: dict) -> tuple:
    """Return next two available xl/tables/tableN.xml paths."""
    nums = []
    for name in parts:
        m = re.match(r"xl/tables/table(\d+)\.xml$", name)
        if m:
            nums.append(int(m.group(1)))
    base = (max(nums) + 1) if nums else 1
    return f"xl/tables/table{base}.xml", f"xl/tables/table{base + 1}.xml"


def _max_table_id(parts: dict) -> int:
    """Highest id= value across all table XML files (workbook-unique constraint)."""
    ids = []
    for name, data in parts.items():
        if re.match(r"xl/tables/table\d+\.xml$", name):
            m = re.search(rb"<table\b[^>]+\bid=\"(\d+)\"", data)
            if m:
                ids.append(int(m.group(1)))
    return max(ids) if ids else 0


def _max_rel_id(wb_rels: bytes) -> int:
    nums = [int(x) for x in re.findall(rb'\bId="rId(\d+)"', wb_rels)]
    return max(nums) if nums else 0


def _max_sheet_id(wb_xml: bytes) -> int:
    nums = [int(x) for x in re.findall(rb'\bsheetId="(\d+)"', wb_xml)]
    return max(nums) if nums else 0


def _strip_xr_uids(data: bytes) -> bytes:
    """
    Remove xr*:uid="{GUID}" revision-tracking attributes.
    These are Excel-proprietary; duplicates cause co-authoring conflicts.
    Stripping them is safe: Excel treats them as optional annotations.
    """
    return re.sub(rb'\s+xr\d*:uid="\{[0-9A-Fa-f-]+\}"', b"", data)


# ---------------------------------------------------------------------------
# Step 3 — clone sheet XML
# ---------------------------------------------------------------------------

def _clone_sheet_xml(data: bytes, old_table_name: str, new_table_name: str) -> bytes:
    """
    Patch the cloned sheet XML:
      (a) rename all structured references  e.g. Saving_Table2410[ → Saving_Table2410_2[
      (b) deactivate the sheet's tab selection
      (c) remove <drawing> and <legacyDrawing> elements (charts stay on original)
      (d) strip xr:uid revision attributes
    """
    # (a) Structured reference rename — 13,883 occurrences verified in pre-flight
    #     Only 'TableName[' pattern used; bare name never appears in sheet XML.
    data = data.replace(
        (old_table_name + "[").encode(),
        (new_table_name + "[").encode(),
    )

    # (b) Deactivate tab so both sheets are not selected simultaneously on open
    data = data.replace(b'tabSelected="1"', b'tabSelected="0"')

    # (c) Charts are hardcoded to "Saving_Calculations!" data sources; removing
    #     the drawing reference prevents a broken drawing on the clone.
    #     Pattern: <drawing r:id="..."/>  and  <legacyDrawing r:id="..."/>
    data = re.sub(rb"<drawing\b[^>]*/>", b"", data)
    data = re.sub(rb"<legacyDrawing\b[^>]*/>", b"", data)

    # (d) Revision tracking GUIDs
    data = _strip_xr_uids(data)

    return data


# ---------------------------------------------------------------------------
# Step 4 — clone sheet _rels
# ---------------------------------------------------------------------------

def _zip_to_rel_target(zip_path: str) -> str:
    """
    Convert a zip part path to the relative target used in a worksheet _rels.
    "xl/tables/table1.xml"  →  "../tables/table1.xml"
    """
    assert zip_path.startswith("xl/"), f"Expected xl/ prefix: {zip_path}"
    return "../" + zip_path[3:]  # strip leading "xl/"


def _clone_sheet_rels(
    data: bytes,
    old_main_zip: str,
    new_main_zip: str,
    old_summ_zip: str,
    new_summ_zip: str,
) -> bytes:
    """
    In the cloned sheet's _rels:
      - Retarget table relationships to the new table parts
      - Remove drawing and vmlDrawing relationships (chart POC scope decision)
    """
    # Retarget main table
    data = data.replace(
        _zip_to_rel_target(old_main_zip).encode(),
        _zip_to_rel_target(new_main_zip).encode(),
    )
    # Retarget summary table
    data = data.replace(
        _zip_to_rel_target(old_summ_zip).encode(),
        _zip_to_rel_target(new_summ_zip).encode(),
    )

    # Remove drawing-type relationships so the cloned sheet has no dangling refs
    # Matches both REL_DRAWING and REL_VMLDRAWING by suffix
    data = re.sub(
        rb"<Relationship\b[^>]+Type=\"[^\"]*(?:drawing|vmlDrawing)\"[^>]*/>",
        b"",
        data,
    )

    return data


# ---------------------------------------------------------------------------
# Step 5 — clone main table XML
# ---------------------------------------------------------------------------

def _clone_main_table_xml(
    data: bytes,
    old_name: str,
    new_name: str,
    old_id: int,
    new_id: int,
) -> bytes:
    """
    Patch the cloned main table XML:
      - Rename name= and displayName= on the root <table> element
      - Update the workbook-unique id= attribute
      - Rename all 13 calculatedColumnFormula structured references
      - Strip xr:uid attributes
    """
    # Root element: name="Saving_Table2410" displayName="Saving_Table2410"
    # These two attributes appear adjacently on the root <table> element.
    # Pre-flight confirmed this exact substring exists exactly once.
    data = data.replace(
        f'name="{old_name}" displayName="{old_name}"'.encode(),
        f'name="{new_name}" displayName="{new_name}"'.encode(),
    )

    # Table id= — root element appears before any tableColumn in the file.
    # count=1 ensures only the root element's id is replaced.
    # Pre-flight: no tableColumn has id=9; root id at byte 375 vs first col at 859.
    data = data.replace(
        f'id="{old_id}"'.encode(),
        f'id="{new_id}"'.encode(),
        1,
    )

    # calculatedColumnFormula elements (13 occurrences)
    data = data.replace(
        (old_name + "[").encode(),
        (new_name + "[").encode(),
    )

    data = _strip_xr_uids(data)
    return data


# ---------------------------------------------------------------------------
# Step 6 — clone summary table XML
# ---------------------------------------------------------------------------

def _clone_summ_table_xml(
    data: bytes,
    old_name: str,
    new_name: str,
    old_id: int,
    new_id: int,
) -> bytes:
    """
    Patch the cloned summary table (FiveYearPlan / Table43611).
    No structured references — only root element rename and id update.
    """
    data = data.replace(
        f'name="{old_name}" displayName="{old_name}"'.encode(),
        f'name="{new_name}" displayName="{new_name}"'.encode(),
    )
    data = data.replace(
        f'id="{old_id}"'.encode(),
        f'id="{new_id}"'.encode(),
        1,
    )
    data = _strip_xr_uids(data)
    return data


# ---------------------------------------------------------------------------
# Step 7 — update workbook.xml
# ---------------------------------------------------------------------------

def _update_workbook_xml(
    data: bytes,
    clone_display_name: str,
    new_sheet_id: int,
    new_rel_id: str,
) -> bytes:
    """Append a new <sheet> entry inside the existing <sheets> block."""
    new_element = (
        f'<sheet name="{xml_escape(clone_display_name)}" '
        f'sheetId="{new_sheet_id}" '
        f'r:id="{new_rel_id}"/>'
    ).encode()
    return data.replace(b"</sheets>", new_element + b"</sheets>")


# ---------------------------------------------------------------------------
# Step 8 — update workbook.xml.rels
# ---------------------------------------------------------------------------

def _update_workbook_rels(
    data: bytes, new_rel_id: str, new_sheet_part: str
) -> bytes:
    """Append a new worksheet <Relationship> to workbook.xml.rels."""
    # Target is relative from xl/: strip the "xl/" prefix
    target = new_sheet_part[3:]    # "xl/worksheets/sheetN.xml" → "worksheets/sheetN.xml"
    new_rel = (
        f'<Relationship Id="{new_rel_id}" '
        f'Type="{REL_WORKSHEET}" '
        f'Target="{target}"/>'
    ).encode()
    return data.replace(b"</Relationships>", new_rel + b"</Relationships>")


# ---------------------------------------------------------------------------
# Step 9 — update [Content_Types].xml
# ---------------------------------------------------------------------------

def _update_content_types(
    data: bytes,
    new_sheet_part: str,
    new_table_main: str,
    new_table_summ: str,
) -> bytes:
    """
    Register the three new parts with their content types.
    Missing Override entries cause Excel to show "We found a problem…" on open.
    """
    def _override(part: str, ct: str) -> bytes:
        return f'<Override PartName="/{part}" ContentType="{ct}"/>'.encode()

    inserts = (
        _override(new_sheet_part, CT_WORKSHEET)
        + _override(new_table_main, CT_TABLE)
        + _override(new_table_summ, CT_TABLE)
    )
    return data.replace(b"</Types>", inserts + b"</Types>")


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------

def clone_saving_calculations(
    template_path: str,
    output_path: str,
    clone_display_name: str = "Saving_Calculations_2",
    clone_suffix: str = CLONE_SUFFIX,
) -> dict:
    """
    Clone the Saving_Calculations sheet from template_path.
    Writes output to output_path.
    clone_suffix controls the table-name suffix (default "_2").
    Pass "_3", "_4" etc. when calling on an already-cloned workbook.
    Returns a diagnostics dict consumed by _print_report().
    """
    diag = {}

    # ── Load ────────────────────────────────────────────────────────────────
    parts = _load_parts(template_path)
    diag["parts_loaded"] = len(parts)

    # ── Discover source parts ────────────────────────────────────────────────
    sheet_path, sheet_rels_path, sheet_display_name = _find_sheet_part(parts)
    diag["source_sheet_part"] = sheet_path
    diag["source_sheet_name"] = sheet_display_name
    diag["source_sheet_rels"] = sheet_rels_path

    (main_zip, main_name), (summ_zip, summ_name) = _find_table_parts(
        parts, sheet_rels_path, sheet_path
    )
    diag["main_table_part"] = main_zip
    diag["main_table_name"] = main_name
    diag["summ_table_part"] = summ_zip
    diag["summ_table_name"] = summ_name

    # Extract current IDs from source tables
    main_id = int(
        re.search(rb"<table\b[^>]+\bid=\"(\d+)\"", parts[main_zip]).group(1)
    )
    summ_id = int(
        re.search(rb"<table\b[^>]+\bid=\"(\d+)\"", parts[summ_zip]).group(1)
    )
    diag["main_table_id_src"] = main_id
    diag["summ_table_id_src"] = summ_id

    # ── Allocate new part paths and IDs ──────────────────────────────────────
    new_sheet_part = _next_sheet_part(parts)
    new_table_main_part, new_table_summ_part = _next_table_parts(parts)
    new_main_id   = _max_table_id(parts) + 1
    new_summ_id   = new_main_id + 1
    new_sheet_id  = _max_sheet_id(parts["xl/workbook.xml"]) + 1
    new_rel_id    = f"rId{_max_rel_id(parts['xl/_rels/workbook.xml.rels']) + 1}"
    new_main_name = main_name + clone_suffix
    new_summ_name = summ_name + clone_suffix

    diag.update(
        {
            "new_sheet_part": new_sheet_part,
            "new_table_main_part": new_table_main_part,
            "new_table_summ_part": new_table_summ_part,
            "new_main_table_name": new_main_name,
            "new_summ_table_name": new_summ_name,
            "new_main_table_id": new_main_id,
            "new_summ_table_id": new_summ_id,
            "new_sheet_id": new_sheet_id,
            "new_rel_id": new_rel_id,
        }
    )

    # ── Steps 3–6: produce cloned XML bytes ──────────────────────────────────
    cloned_sheet = _clone_sheet_xml(parts[sheet_path], main_name, new_main_name)
    cloned_rels = _clone_sheet_rels(
        parts[sheet_rels_path],
        main_zip, new_table_main_part,
        summ_zip, new_table_summ_part,
    )
    cloned_main_table = _clone_main_table_xml(
        parts[main_zip], main_name, new_main_name, main_id, new_main_id
    )
    cloned_summ_table = _clone_summ_table_xml(
        parts[summ_zip], summ_name, new_summ_name, summ_id, new_summ_id
    )

    # Substitution count verification
    diag["sr_in_source"]  = parts[sheet_path].count((main_name + "[").encode())
    diag["sr_in_clone"]   = cloned_sheet.count((new_main_name + "[").encode())
    diag["sr_old_in_clone"] = cloned_sheet.count((main_name + "[").encode())

    # ── Steps 7–9: update manifest files in-place ────────────────────────────
    parts["xl/workbook.xml"] = _update_workbook_xml(
        parts["xl/workbook.xml"], clone_display_name, new_sheet_id, new_rel_id
    )
    parts["xl/_rels/workbook.xml.rels"] = _update_workbook_rels(
        parts["xl/_rels/workbook.xml.rels"], new_rel_id, new_sheet_part
    )
    parts["[Content_Types].xml"] = _update_content_types(
        parts["[Content_Types].xml"],
        new_sheet_part, new_table_main_part, new_table_summ_part,
    )

    # ── Step 10: delete calcChain ─────────────────────────────────────────────
    diag["calc_chain_present"] = "xl/calcChain.xml" in parts
    parts.pop("xl/calcChain.xml", None)

    # ── Register new parts ────────────────────────────────────────────────────
    new_sheet_base = os.path.basename(new_sheet_part)       # "sheet4.xml"
    new_sheet_rels_part = f"xl/worksheets/_rels/{new_sheet_base}.rels"

    parts[new_sheet_part]       = cloned_sheet
    parts[new_sheet_rels_part]  = cloned_rels
    parts[new_table_main_part]  = cloned_main_table
    parts[new_table_summ_part]  = cloned_summ_table

    diag["new_sheet_rels_part"] = new_sheet_rels_part

    # ── Write output ZIP ─────────────────────────────────────────────────────
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_bytes(buf.getvalue())

    diag["output_path"]       = output_path
    diag["output_size_bytes"] = len(buf.getvalue())
    diag["parts_in_output"]   = len(parts)

    return diag


# ---------------------------------------------------------------------------
# Diagnostic report
# ---------------------------------------------------------------------------

def _print_report(diag: dict) -> None:
    SEP = "=" * 62
    print(f"\n{SEP}")
    print("ZIP-LEVEL CLONE — DIAGNOSTIC REPORT")
    print(SEP)

    print("\n[DISCOVERY]")
    print(f"  Source sheet part      : {diag['source_sheet_part']}")
    print(f"  Source sheet name      : {diag['source_sheet_name']!r}")
    print(f"  Main table             : {diag['main_table_name']!r}")
    print(f"    part                 : {diag['main_table_part']}")
    print(f"    id (source)          : {diag['main_table_id_src']}")
    print(f"  Summary table          : {diag['summ_table_name']!r}")
    print(f"    part                 : {diag['summ_table_part']}")
    print(f"    id (source)          : {diag['summ_table_id_src']}")

    print("\n[ALLOCATION]")
    print(f"  New sheet part         : {diag['new_sheet_part']}")
    print(f"  New sheet rels         : {diag['new_sheet_rels_part']}")
    print(f"  New main table part    : {diag['new_table_main_part']}")
    print(f"  New summ table part    : {diag['new_table_summ_part']}")
    print(f"  New main table name    : {diag['new_main_table_name']!r}")
    print(f"  New summ table name    : {diag['new_summ_table_name']!r}")
    print(f"  New main table id      : {diag['new_main_table_id']}")
    print(f"  New summ table id      : {diag['new_summ_table_id']}")
    print(f"  New workbook sheetId   : {diag['new_sheet_id']}")
    print(f"  New workbook rId       : {diag['new_rel_id']}")

    print("\n[SUBSTITUTION CHECK]")
    src  = diag["sr_in_source"]
    cln  = diag["sr_in_clone"]
    old  = diag["sr_old_in_clone"]
    ok_a = "PASS" if src == cln else "FAIL"
    ok_b = "PASS" if old == 0   else "FAIL"
    print(f"  Old name '[' in source       : {src}")
    print(f"  New name '[' in clone        : {cln}  [{ok_a}]")
    print(f"  Old name '[' still in clone  : {old}  [{ok_b}]")

    print("\n[MANIFEST]")
    print(f"  calcChain.xml deleted  : {diag['calc_chain_present']}")
    print(f"  Parts in source        : {diag['parts_loaded']}")
    print(f"  Parts in output        : {diag['parts_in_output']}")

    print("\n[OUTPUT]")
    print(f"  Path  : {diag['output_path']}")
    print(f"  Size  : {diag['output_size_bytes']:,} bytes")
    print()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    template_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_TEMPLATE
    output_path   = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUTPUT

    print(f"Template : {template_path}")
    print(f"Output   : {output_path}")

    diag = clone_saving_calculations(template_path, output_path)
    _print_report(diag)
    print("Done.")
