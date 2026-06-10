#!/usr/bin/env python3
"""
test_chart_preservation.py
--------------------------
Validates that chart parts are correctly cloned when Saving_Calculations
is duplicated via clone_saving_calculations().

Checks:
  A — structural: drawing4.xml, drawing4.xml.rels, chart8–14.xml, chart8–14.xml.rels
      all exist in the output ZIP.
  B — content-types: [Content_Types].xml has Override entries for drawing4
      and each of chart8–14.
  C — sheet wiring: cloned sheetN.xml has <drawing r:id="rId2"/>;
      sheetN.xml.rels has a drawing relationship targeting drawing4.xml.
  D — formula correctness: cloned charts contain new sheet name,
      NOT the old sheet name (except chart6-equiv which uses an external ref).
  E — original untouched: chart1–7.xml still reference Saving_Calculations!
  F — drawing rels: drawing4.xml.rels has exactly 7 chart + 2 image relationships.
  G — part count: output has 19 more parts than input (1 drawing + 1 drawing rels
      + 7 charts + 7 chart rels + 1 sheet + 1 sheet rels + 2 tables = 20; minus
      calcChain = 19).

Prints PASS/FAIL for each check, exits non-zero on any failure.
"""

import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(__file__))
from excel_clone_poc import clone_saving_calculations

TEMPLATE = os.path.join(
    os.path.dirname(__file__),
    "excel_templates",
    "1_0_region_site_name_saving_calculations.xlsx",
)
OUTPUT = os.path.join(os.path.dirname(__file__), "test_chart_preservation_output.xlsx")

PASS_LIST = []
FAIL_LIST = []


def check(label: str, cond: bool, detail: str = "") -> None:
    if cond:
        PASS_LIST.append(label)
        print(f"  PASS  {label}")
    else:
        FAIL_LIST.append(f"{label}  {detail}")
        print(f"  FAIL  {label}  {detail}")


def run() -> None:
    # ── Generate ─────────────────────────────────────────────────────────────
    print(f"\nGenerating: {OUTPUT}")
    diag = clone_saving_calculations(TEMPLATE, OUTPUT, "Saving_Calculations_2", "_2")
    print(f"  Parts in output: {diag['parts_in_output']}")
    print(f"  Chart formula patches: {diag['chart_formula_replacements']}")

    # ── Load output ZIP ───────────────────────────────────────────────────────
    with zipfile.ZipFile(OUTPUT) as zf:
        names = set(zf.namelist())
        parts = {n: zf.read(n) for n in names}

    # ── A: Structural presence ────────────────────────────────────────────────
    print("\n── A: Structural presence ───────────────────────────────────────")
    check("A01: drawing4.xml exists",
          "xl/drawings/drawing4.xml" in parts)
    check("A02: drawing4.xml.rels exists",
          "xl/drawings/_rels/drawing4.xml.rels" in parts)
    for i in range(7):
        n = 8 + i
        check(f"A{i+3:02d}: chart{n}.xml exists",
              f"xl/charts/chart{n}.xml" in parts)
    for i in range(7):
        n = 8 + i
        check(f"A{i+10:02d}: chart{n}.xml.rels exists",
              f"xl/charts/_rels/chart{n}.xml.rels" in parts)

    # ── B: Content-Types registrations ───────────────────────────────────────
    print("\n── B: Content-Types registrations ───────────────────────────────")
    ct = parts["[Content_Types].xml"].decode()
    check("B01: drawing4 Override in Content-Types",
          'PartName="/xl/drawings/drawing4.xml"' in ct)
    for i in range(7):
        n = 8 + i
        check(f"B{i+2:02d}: chart{n} Override in Content-Types",
              f'PartName="/xl/charts/chart{n}.xml"' in ct)

    # ── C: Sheet wiring ───────────────────────────────────────────────────────
    print("\n── C: Sheet wiring ──────────────────────────────────────────────")
    # Find the cloned sheet XML (sheet4.xml when template has 3 sheets)
    clone_sheet_path = diag["new_sheet_part"]
    clone_rels_path  = diag["new_sheet_rels_part"]

    clone_sheet_xml  = parts[clone_sheet_path].decode()
    clone_rels_xml   = parts[clone_rels_path].decode()

    check("C01: <drawing r:id=\"rId2\"/> present in cloned sheet XML",
          'r:id="rId2"' in clone_sheet_xml and "<drawing" in clone_sheet_xml,
          f"drawing element not found in {clone_sheet_path}")

    check("C02: drawing relationship present in cloned sheet rels",
          "relationships/drawing" in clone_rels_xml,
          f"no drawing rel in {clone_rels_path}")

    check("C03: drawing relationship targets drawing4.xml",
          "drawing4.xml" in clone_rels_xml,
          f"drawing4.xml not referenced in {clone_rels_path}")

    check("C04: vmlDrawing stripped from cloned sheet rels",
          "vmlDrawing" not in clone_rels_xml)

    # ── D: Formula correctness in cloned charts ───────────────────────────────
    print("\n── D: Formula correctness in cloned charts ──────────────────────")
    # chart6 equivalent (chart13) uses external [1]Sheet1! — skipped from name check
    external_chart_idx = 5   # chart6 is index 5 (0-based) among the 7
    for i in range(7):
        n = 8 + i
        data = parts[f"xl/charts/chart{n}.xml"].decode()
        formulas = re.findall(r"<c:f>[^<]+</c:f>", data)
        if i == external_chart_idx:
            # chart6 equivalent: external ref only, no Saving_Calculations refs
            check(f"D{i+1:02d}: chart{n} has no Saving_Calculations_2 ref (external chart)",
                  "Saving_Calculations_2!" not in data and "Saving_Calculations!" not in data,
                  f"unexpected sheet ref found")
        else:
            check(f"D{i+1:02d}: chart{n} refs new sheet name",
                  "Saving_Calculations_2!" in data,
                  f"Saving_Calculations_2! not found — formulas: {formulas[:2]}")
            check(f"D{i+1:02d}b: chart{n} has no old sheet name",
                  "Saving_Calculations!" not in data,
                  f"old sheet name still present")

    # ── E: Original charts untouched ─────────────────────────────────────────
    print("\n── E: Original charts untouched ─────────────────────────────────")
    for i in range(7):
        n = i + 1
        data = parts[f"xl/charts/chart{n}.xml"].decode()
        if n == 6:   # external chart — no sheet ref expected
            continue
        check(f"E{i+1:02d}: chart{n} still refs Saving_Calculations!",
              "Saving_Calculations!" in data,
              f"original chart{n} formula refs missing")

    # ── F: Drawing rels structure ─────────────────────────────────────────────
    print("\n── F: Drawing4 rels structure ───────────────────────────────────")
    d4_rels = parts["xl/drawings/_rels/drawing4.xml.rels"].decode()
    chart_rels_count = len(re.findall(r'Type="[^"]*?/chart"', d4_rels))
    image_rels_count = len(re.findall(r'Type="[^"]*?/image"', d4_rels))
    check("F01: drawing4.xml.rels has 7 chart relationships",
          chart_rels_count == 7, f"found {chart_rels_count}")
    check("F02: drawing4.xml.rels has 2 image relationships",
          image_rels_count == 2, f"found {image_rels_count}")
    check("F03: drawing4.xml.rels targets chart8 through chart14",
          all(f"chart{8+i}.xml" in d4_rels for i in range(7)))
    check("F04: drawing4.xml.rels targets shared image3.png",
          "image3.png" in d4_rels)
    check("F05: drawing4.xml.rels targets shared image4.png",
          "image4.png" in d4_rels)

    # ── G: Part count ─────────────────────────────────────────────────────────
    print("\n── G: Part count ────────────────────────────────────────────────")
    # Template: 80 parts.  Additions: +1 sheet, +1 sheet rels, +2 tables,
    #   +1 drawing, +1 drawing rels, +7 charts, +7 chart rels = +20
    # Removals: -1 calcChain = -1.  Net: +19 → expected 99
    check("G01: output has 99 parts (80 + 20 additions - 1 calcChain)",
          diag["parts_in_output"] == 99,
          f"found {diag['parts_in_output']}")

    # ── H: cNvPr id uniqueness (root cause of Excel repair) ──────────────────
    print("\n── H: cNvPr id uniqueness ───────────────────────────────────────")
    def _cnvpr_ids(xml_bytes):
        return [int(x) for x in re.findall(rb'<[a-zA-Z:]+cNvPr\s+id="(\d+)"', xml_bytes)]

    ids_d1 = _cnvpr_ids(parts["xl/drawings/drawing1.xml"])
    ids_d4 = _cnvpr_ids(parts["xl/drawings/drawing4.xml"])

    check("H01: drawing4 has 13 cNvPr ids (same shape count as drawing1)",
          len(ids_d4) == 13, f"found {len(ids_d4)}")

    overlap = set(ids_d1) & set(ids_d4)
    check("H02: no cNvPr id overlap between drawing1 and drawing4",
          len(overlap) == 0, f"colliding ids: {sorted(overlap)}")

    check("H03: all drawing4 cNvPr ids > drawing1 max (16)",
          all(i > 16 for i in ids_d4), f"ids: {sorted(ids_d4)}")

    check("H04: drawing4 cNvPr ids are sequential from 17",
          ids_d4 == list(range(17, 30)), f"expected 17–29, got {ids_d4}")

    # drawing2.xml and drawing3.xml are chartshapes parts (pre-existing template
    # collision between them); we only need drawing4 to be clean vs all others.
    ids_others = (
        _cnvpr_ids(parts["xl/drawings/drawing1.xml"]) +
        _cnvpr_ids(parts["xl/drawings/drawing2.xml"]) +
        _cnvpr_ids(parts["xl/drawings/drawing3.xml"])
    )
    overlap_d4 = set(ids_d4) & set(ids_others)
    check("H05: drawing4 cNvPr ids don't collide with drawing1/2/3",
          len(overlap_d4) == 0, f"colliding ids: {sorted(overlap_d4)}")

    # ── I: a16:creationId GUID uniqueness ─────────────────────────────────────
    print("\n── I: a16:creationId GUID uniqueness ────────────────────────────")
    # Count by searching for the element name directly; avoids regex issues with
    # URLs (containing '/') that appear in the xmlns attribute of these elements.
    def _count_creation_ids(xml_bytes):
        return xml_bytes.count(b"a16:creationId")

    n_d1 = _count_creation_ids(parts["xl/drawings/drawing1.xml"])
    n_d4 = _count_creation_ids(parts["xl/drawings/drawing4.xml"])

    check("I01: drawing4 has no a16:creationId elements (stripped)",
          n_d4 == 0, f"found {n_d4} a16:creationId occurrences")

    check("I02: drawing1 original creationIds untouched (13 present)",
          n_d1 == 13, f"found {n_d1}")

    # ── Report ────────────────────────────────────────────────────────────────
    SEP = "=" * 60
    print(f"\n{SEP}")
    print("CHART PRESERVATION — VALIDATION REPORT")
    print(SEP)
    total = len(PASS_LIST) + len(FAIL_LIST)
    print(f"\n  PASS: {len(PASS_LIST)}/{total}   FAIL: {len(FAIL_LIST)}/{total}")
    if FAIL_LIST:
        print("\n  FAILURES:")
        for msg in FAIL_LIST:
            print(f"    FAIL  {msg}")
    print()
    print("RESULT: PASS" if not FAIL_LIST else "RESULT: FAIL")
    print()


if __name__ == "__main__":
    run()
    sys.exit(0 if not FAIL_LIST else 1)
