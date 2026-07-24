#!/usr/bin/env python3
"""
Phase 0: Tokenize ABB Energy Appraisal Standard report templates.

Replaces Spain-specific data values with {{TOKEN}} placeholders so that
the report generator can do simple string substitution instead of
fragile XML-node-position arithmetic.

Run once after any template update. Backs up originals before modifying.

Usage:
    python3 tools/prepare_templates.py
"""

import zipfile, re, io, os, shutil
from pathlib import Path

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / 'report_templates'

FILES = {
    'executive': 'ea_report_template_standard_top 10.docx',
    'complete':  'ea_report_template_standard_all_assets.docx',
}

# ── XML structure helpers ─────────────────────────────────────────────────────

def find_tables(xml):
    """Return list of (start, end) for every top-level <w:tbl> in xml."""
    tables, pos = [], 0
    while True:
        s = xml.find('<w:tbl>', pos)
        if s == -1:
            break
        d, i = 1, s + 7
        while i < len(xml):
            if xml[i:i+7] == '<w:tbl>':
                d += 1; i += 7
            elif xml[i:i+8] == '</w:tbl>':
                d -= 1; i += 8
                if d == 0:
                    break
            else:
                i += 1
        tables.append((s, i))
        pos = i
    return tables


def find_trs(xml, tbl_start, tbl_end):
    """Return list of (abs_start, abs_end) for every <w:tr> within table bounds."""
    segment = xml[tbl_start:tbl_end]
    return [
        (tbl_start + m.start(), tbl_start + m.end())
        for m in re.finditer(r'<w:tr[ >].*?</w:tr>', segment, re.DOTALL)
    ]


def find_tcs(tr_xml):
    """Return list of (start, end) for every <w:tc> in a TR string."""
    tcs, pos = [], 0
    while True:
        s = tr_xml.find('<w:tc>', pos)
        if s == -1:
            break
        d, i = 1, s + 6
        while i < len(tr_xml):
            if tr_xml[i:i+6] == '<w:tc>':
                d += 1; i += 6
            elif tr_xml[i:i+7] == '</w:tc>':
                d -= 1; i += 7
                if d == 0:
                    break
            else:
                i += 1
        tcs.append((s, i))
        pos = i
    return tcs


def _first(pattern, text):
    m = re.search(pattern, text, re.DOTALL)
    return m.group(1) if m else ''

def get_tcpr(tc_xml):  return _first(r'(<w:tcPr>.*?</w:tcPr>)', tc_xml)
def get_ppr(xml_frag): return _first(r'(<w:pPr>.*?</w:pPr>)',   xml_frag)
def get_rpr(xml_frag): return _first(r'(<w:rPr>.*?</w:rPr>)',   xml_frag)


def build_single_run_cell(tc_xml, token):
    """
    Replace all paragraph content in tc_xml with one run containing token.
    Preserves <w:tcPr>, the first <w:pPr>, and the first <w:rPr>.
    """
    tcpr = get_tcpr(tc_xml)
    ppr  = get_ppr(tc_xml)
    rpr  = get_rpr(tc_xml)
    space = ' xml:space="preserve"' if ' ' in token else ''
    new_p = f'<w:p>{ppr}<w:r>{rpr}<w:t{space}>{token}</w:t></w:r></w:p>'
    return f'<w:tc>{tcpr}{new_p}</w:tc>'


# ── Targeted <w:t> replacement within a bounded XML region ───────────────────

def repl_wt(xml, region_s, region_e, old_val, new_val, count=1):
    """
    Within xml[region_s:region_e], replace every <w:t ...>old_val</w:t>
    with <w:t ...>new_val</w:t>.  count=0 means replace all occurrences.
    Returns modified full xml string.
    """
    region = xml[region_s:region_e]
    pat = r'(<w:t(?:[^>]*)>)' + re.escape(old_val) + r'(</w:t>)'
    new_region = re.sub(pat, r'\g<1>' + new_val + r'\g<2>', region, count=count)
    return xml[:region_s] + new_region + xml[region_e:]


# ── Cover table helpers ───────────────────────────────────────────────────────

def replace_cover_value_cell(xml, label_text, token):
    """
    In the first table (cover info block), find the TR whose first cell
    text contains label_text.  Replace the second cell entirely with token.
    label_text should be specific enough to identify one row (e.g. '>Customer<').
    """
    tables = find_tables(xml)
    if not tables:
        return xml
    tbl_s, tbl_e = tables[0]
    for tr_s, tr_e in find_trs(xml, tbl_s, tbl_e):
        tr_xml = xml[tr_s:tr_e]
        if label_text not in tr_xml:
            continue
        tcs = find_tcs(tr_xml)
        if len(tcs) < 2:
            continue
        tc_s, tc_e = tcs[1]
        old_tc = tr_xml[tc_s:tc_e]
        new_tc = build_single_run_cell(old_tc, token)
        new_tr = tr_xml[:tc_s] + new_tc + tr_xml[tc_e:]
        xml = xml[:tr_s] + new_tr + xml[tr_e:]
        break
    return xml


# ── Table / row / cell locators ──────────────────────────────────────────────

def table_cell_bounds(xml, tbl_idx, tr_idx, cell_idx):
    """Return (abs_start, abs_end) of a cell by table/row/cell index."""
    tables = find_tables(xml)
    if tbl_idx >= len(tables):
        return None, None
    tbl_s, tbl_e = tables[tbl_idx]
    trs = find_trs(xml, tbl_s, tbl_e)
    if tr_idx >= len(trs):
        return None, None
    tr_s, tr_e = trs[tr_idx]
    tr_xml = xml[tr_s:tr_e]
    tcs = find_tcs(tr_xml)
    if cell_idx >= len(tcs):
        return None, None
    tc_s, tc_e = tcs[cell_idx]
    return tr_s + tc_s, tr_s + tc_e


def summary_cell_bounds(xml, tr_idx, cell_idx):
    """
    Return (abs_start, abs_end) of cell_idx-th <w:tc> in tr_idx-th row
    of the summary table (always Table index 2 in both new templates).
    """
    return table_cell_bounds(xml, 2, tr_idx, cell_idx)


def replace_table_cell(xml, tbl_idx, tr_idx, cell_idx, token):
    """Replace entire paragraph content of one cell with a single-run token."""
    cs, ce = table_cell_bounds(xml, tbl_idx, tr_idx, cell_idx)
    if cs is None:
        return xml
    new_tc = build_single_run_cell(xml[cs:ce], token)
    return xml[:cs] + new_tc + xml[ce:]


# ── ROW_TEMPLATE marker insertion ─────────────────────────────────────────────

def insert_row_template(xml, table_idx):
    """
    Replace the first cell of TR[1] (first data row) in the specified
    table with {{ROW_TEMPLATE}}.  TR[0] is always the column-header row.
    """
    tables = find_tables(xml)
    if table_idx >= len(tables):
        return xml
    tbl_s, tbl_e = tables[table_idx]
    trs = find_trs(xml, tbl_s, tbl_e)
    if len(trs) < 2:
        return xml
    tr_s, tr_e = trs[1]
    tr_xml = xml[tr_s:tr_e]
    tcs = find_tcs(tr_xml)
    if not tcs:
        return xml
    tc_s, tc_e = tcs[0]
    old_tc = tr_xml[tc_s:tc_e]
    new_tc = build_single_run_cell(old_tc, '{{ROW_TEMPLATE}}')
    new_tr = tr_xml[:tc_s] + new_tc + tr_xml[tc_e:]
    return xml[:tr_s] + new_tr + xml[tr_e:]


# ── Footer tokenization ───────────────────────────────────────────────────────

def tokenize_footer(xml):
    """Replace ISO date in footer files."""
    return repl_wt(xml, 0, len(xml), '2026-06-23', '{{REPORT_DATE_ISO}}', count=0)


# ── Executive template tokenization ──────────────────────────────────────────

def tokenize_executive(xml):

    # ── Cover table ──────────────────────────────────────────────────────────
    # All cover value cells in Executive are single (unfragmented) nodes.
    # build_single_run_cell handles them uniformly.
    for label, token in [
        ('>Customer<',        '{{CUSTOMER}}'),
        ('>Plants<',          '{{PLANT}}'),
        ('>Data Source<',     '{{DATA_SOURCE}}'),
        ('>Date of Report<',  '{{REPORT_DATE}}'),
        ('># of Assets<',     '{{NPV_COUNT}} / {{TOTAL_ASSETS}}'),
        ('>Electricity Cost<','{{ELEC_PRICE}} {{CURRENCY}}/kWh'),
        ('>Carbon Intensity<','{{CO2_INTENSITY}} kg CO2/kWh'),
    ]:
        xml = replace_cover_value_cell(xml, label, token)

    # ── Summary TR[0] ─────────────────────────────────────────────────────────
    # Cell[0]: EUR [TOP10_SAVINGS] label EUR [NPV_SAVINGS(fragmented)] label
    #   Fragments: '26,529' (single)  |  ' 3'+'1,229' (split)
    cs, ce = summary_cell_bounds(xml, 0, 0)
    xml = repl_wt(xml, cs, ce, 'EUR',    '{{CURRENCY}}', count=0)
    xml = repl_wt(xml, cs, ce, '26,529', '{{TOP10_ANNUAL_SAVINGS}}')
    xml = repl_wt(xml, cs, ce, ' 3',     ' {{ANNUAL_SAVINGS}}')   # keep leading space
    xml = repl_wt(xml, cs, ce, '1,229',  '')                       # clear fragment

    # Cell[2]: second <w:p> contains NPV count/total:  '2'+'1/'+'45' 'motors ' 'NPV positive'
    cs, ce = summary_cell_bounds(xml, 0, 2)
    xml = repl_wt(xml, cs, ce, '2',  '{{NPV_COUNT}} / ')  # absorbs the fused separator
    xml = repl_wt(xml, cs, ce, '1/', '')                   # clear second digit + slash
    xml = repl_wt(xml, cs, ce, '45', '{{TOTAL_ASSETS}}')

    # ── Summary TR[2] ─────────────────────────────────────────────────────────
    # Cell[0]: EUR [TOP10_INVEST] label  EUR [NPV_INVEST] label — all single nodes
    cs, ce = summary_cell_bounds(xml, 2, 0)
    xml = repl_wt(xml, cs, ce, 'EUR',    '{{CURRENCY}}', count=0)
    xml = repl_wt(xml, cs, ce, '49,928', '{{TOP10_INVEST}}')
    xml = repl_wt(xml, cs, ce, '89,528', '{{NPV_INVEST}}')

    # Cell[2]: EUR [TOP10_NPV] label  EUR [NPV_VALUE] label — all single nodes
    cs, ce = summary_cell_bounds(xml, 2, 2)
    xml = repl_wt(xml, cs, ce, 'EUR',    '{{CURRENCY}}', count=0)
    xml = repl_wt(xml, cs, ce, '178,280','{{TOP10_NPV}}')
    xml = repl_wt(xml, cs, ce, '184,638','{{NPV_VALUE}}')

    # ── Summary TR[3] ─────────────────────────────────────────────────────────
    # Payback and IRR live inside TextBox drawings embedded in the cells.
    # The value appears in TWO <w:p> elements (duplicate for visual centering).
    #
    # Cell[0] TextBox: '1'+'.'+'9'+' yrs'  (× 2 paragraphs)
    cs, ce = summary_cell_bounds(xml, 3, 0)
    xml = repl_wt(xml, cs, ce, '1',   '{{TOP10_PAYBACK}}', count=2)
    xml = repl_wt(xml, cs, ce, '.',   '',                  count=2)
    xml = repl_wt(xml, cs, ce, '9',   '',                  count=2)
    # ' yrs' suffix run — static, not replaced

    # Cell[2] TextBox: '42'+'%'  (× 2 paragraphs)
    cs, ce = summary_cell_bounds(xml, 3, 2)
    xml = repl_wt(xml, cs, ce, '42',  '{{IRR_DISPLAY}}', count=2)
    xml = repl_wt(xml, cs, ce, '%',   '',               count=2)

    # ── Summary TR[5] ─────────────────────────────────────────────────────────
    # Cell[0]: '33' ' tCO2' label
    cs, ce = summary_cell_bounds(xml, 5, 0)
    xml = repl_wt(xml, cs, ce, '33', '{{CO2_SAVINGS}}')

    # Cell[1]: '64' ' Vehicles' label
    cs, ce = summary_cell_bounds(xml, 5, 1)
    xml = repl_wt(xml, cs, ce, '64', '{{BEV_COUNT}}')

    # ── ROW_TEMPLATE markers ──────────────────────────────────────────────────
    # Table 3 = Top-10 Energy Savings   (TR[1] = first data row)
    # Table 4 = Top-10 Motor Details
    # Table 5 = All Assets Energy Savings
    # Table 6 = All Assets Motor Details
    for tbl_idx in [3, 4, 5, 6]:
        xml = insert_row_template(xml, tbl_idx)

    # ── Table[3] total rows ───────────────────────────────────────────────────
    # TR[11] = "Total – Top 10" (light-green row)
    for cell_idx, token in [
        (3,  '{{TOTAL_TOP10_ENERGY_CONS}}'),
        (4,  '{{TOTAL_TOP10_ENERGY_COST}}'),
        (5,  '{{TOTAL_TOP10_CO2_CONS}}'),
        (6,  '{{TOTAL_TOP10_SAVINGS_KWH}}'),
        (7,  '{{TOP10_ANNUAL_SAVINGS}}'),
        (8,  '{{TOTAL_TOP10_SAVING_PCT}}'),
        (9,  '{{TOP10_INVEST}}'),
        (10, '{{TOP10_PAYBACK}}'),
        (11, '{{TOTAL_TOP10_CO2_AVOIDED}}'),
    ]:
        xml = replace_table_cell(xml, 3, 11, cell_idx, token)

    # TR[12] = "Total NPV Positive Assets (21)" (dark-green row)
    # Label cell[1] is fragmented: 'Total NPV Positive Assets (' + '2' + '1)'
    cs, ce = table_cell_bounds(xml, 3, 12, 1)
    xml = repl_wt(xml, cs, ce, '2',  '{{NPV_COUNT}}')
    cs, ce = table_cell_bounds(xml, 3, 12, 1)   # recompute after length change
    xml = repl_wt(xml, cs, ce, '1)', ')')
    for cell_idx, token in [
        (3,  '{{TOTAL_NPV_ENERGY_CONS}}'),
        (4,  '{{TOTAL_NPV_ENERGY_COST}}'),
        (5,  '{{TOTAL_NPV_CO2_CONS}}'),
        (6,  '{{TOTAL_NPV_SAVINGS_KWH}}'),
        (7,  '{{ANNUAL_SAVINGS}}'),
        (8,  '{{TOTAL_NPV_SAVING_PCT}}'),
        (9,  '{{NPV_INVEST}}'),
        (10, '{{PAYBACK}}'),
        (11, '{{TOTAL_NPV_CO2_AVOIDED}}'),
    ]:
        xml = replace_table_cell(xml, 3, 12, cell_idx, token)

    # ── Table[5] All Assets total row (mirrors Table[3] TR[12]) ──────────────
    # Table[5] = All Assets Energy Savings; TR[46] = NPV+ total row
    cs, ce = table_cell_bounds(xml, 5, 46, 1)
    xml = repl_wt(xml, cs, ce, '2',  '{{NPV_COUNT}}')
    cs, ce = table_cell_bounds(xml, 5, 46, 1)
    xml = repl_wt(xml, cs, ce, '1)', ')')
    for cell_idx, token in [
        (3,  '{{TOTAL_NPV_ENERGY_CONS}}'),
        (4,  '{{TOTAL_NPV_ENERGY_COST}}'),
        (5,  '{{TOTAL_NPV_CO2_CONS}}'),
        (6,  '{{TOTAL_NPV_SAVINGS_KWH}}'),
        (7,  '{{ANNUAL_SAVINGS}}'),
        (8,  '{{TOTAL_NPV_SAVING_PCT}}'),
        (9,  '{{NPV_INVEST}}'),
        (10, '{{PAYBACK}}'),
        (11, '{{TOTAL_NPV_CO2_AVOIDED}}'),
    ]:
        xml = replace_table_cell(xml, 5, 46, cell_idx, token)

    return xml


# ── Complete Asset template tokenization ─────────────────────────────────────

def tokenize_complete(xml):

    # ── Cover table ──────────────────────────────────────────────────────────
    # TR[4]-TR[7] have heavily fragmented value cells in this template.
    # build_single_run_cell replaces the entire cell content — handles
    # fragmentation by construction (reads rPr from first run, rebuilds as one).
    for label, token in [
        ('>Customer<',        '{{CUSTOMER}}'),
        ('>Plants<',          '{{PLANT}}'),
        ('>Data Source<',     '{{DATA_SOURCE}}'),
        ('>Date of Report<',  '{{REPORT_DATE}}'),
        ('># of Assets<',     '{{NPV_COUNT}} / {{TOTAL_ASSETS}}'),
        ('>Electricity Cost<','{{ELEC_PRICE}} {{CURRENCY}}/kWh'),
        ('>Carbon Intensity<','{{CO2_INTENSITY}} kg CO2/kWh'),
    ]:
        xml = replace_cover_value_cell(xml, label, token)

    # ── Summary TR[0] ─────────────────────────────────────────────────────────
    # Cell[0]: EUR [SAVINGS(heavily fragmented)] label
    #   Fragments: ' 31' + ',' + '2' + '29'  = 31,229
    cs, ce = summary_cell_bounds(xml, 0, 0)
    xml = repl_wt(xml, cs, ce, 'EUR',  '{{CURRENCY}}', count=0)
    xml = repl_wt(xml, cs, ce, ' 31',  ' {{ANNUAL_SAVINGS}}')  # keep leading space
    xml = repl_wt(xml, cs, ce, ',',    '')
    xml = repl_wt(xml, cs, ce, '2',    '')
    xml = repl_wt(xml, cs, ce, '29',   '')

    # Cell[2]: second <w:p>: '21' + '/' + '45' + 'motors ' + 'NPV positive'
    # '/' is already its own clean node — keep it as separator
    cs, ce = summary_cell_bounds(xml, 0, 2)
    xml = repl_wt(xml, cs, ce, '21',  '{{NPV_COUNT}}')
    xml = repl_wt(xml, cs, ce, '45',  '{{TOTAL_ASSETS}}')

    # ── Summary TR[2] ─────────────────────────────────────────────────────────
    # Cell[0]: EUR [INVEST(fragmented)] label
    #   Fragments: '8' + '9,5' + '28'  = 89,528
    cs, ce = summary_cell_bounds(xml, 2, 0)
    xml = repl_wt(xml, cs, ce, 'EUR',  '{{CURRENCY}}', count=0)
    xml = repl_wt(xml, cs, ce, '8',    '{{NPV_INVEST}}')
    xml = repl_wt(xml, cs, ce, '9,5',  '')
    xml = repl_wt(xml, cs, ce, '28',   '')

    # Cell[2]: EUR [NPV(fragmented)] label
    #   Fragments: '184' + ',' + '638'  = 184,638
    cs, ce = summary_cell_bounds(xml, 2, 2)
    xml = repl_wt(xml, cs, ce, 'EUR',  '{{CURRENCY}}', count=0)
    xml = repl_wt(xml, cs, ce, '184',  '{{NPV_VALUE}}')
    xml = repl_wt(xml, cs, ce, ',',    '')
    xml = repl_wt(xml, cs, ce, '638',  '')

    # ── Summary TR[3] ─────────────────────────────────────────────────────────
    # Cell[0] TextBox: '2'+'.'+'9'+' yrs'  (× 2 paragraphs)
    cs, ce = summary_cell_bounds(xml, 3, 0)
    xml = repl_wt(xml, cs, ce, '2',   '{{PAYBACK}}', count=2)
    xml = repl_wt(xml, cs, ce, '.',   '',             count=2)
    xml = repl_wt(xml, cs, ce, '9',   '',             count=2)

    # Cell[2] TextBox: '28'+'%'  (× 2 paragraphs)
    cs, ce = summary_cell_bounds(xml, 3, 2)
    xml = repl_wt(xml, cs, ce, '28',  '{{IRR_DISPLAY}}', count=2)
    xml = repl_wt(xml, cs, ce, '%',   '',                count=2)

    # ── Summary TR[5] ─────────────────────────────────────────────────────────
    cs, ce = summary_cell_bounds(xml, 5, 0)
    xml = repl_wt(xml, cs, ce, '33',  '{{CO2_SAVINGS}}')

    cs, ce = summary_cell_bounds(xml, 5, 1)
    xml = repl_wt(xml, cs, ce, '64',  '{{BEV_COUNT}}')

    # ── ROW_TEMPLATE markers ──────────────────────────────────────────────────
    # Table 3 = All Assets Energy Savings
    # Table 4 = All Assets Motor Details
    for tbl_idx in [3, 4]:
        xml = insert_row_template(xml, tbl_idx)

    # ── Table[3] total row ────────────────────────────────────────────────────
    # TR[46] = "Total NPV Positive Assets (21)" (dark-green row)
    # Label cell[1] is fragmented: 'Total NPV Positive Assets (' + '2' + '1)'
    cs, ce = table_cell_bounds(xml, 3, 46, 1)
    xml = repl_wt(xml, cs, ce, '2',  '{{NPV_COUNT}}')
    cs, ce = table_cell_bounds(xml, 3, 46, 1)   # recompute after length change
    xml = repl_wt(xml, cs, ce, '1)', ')')
    for cell_idx, token in [
        (3,  '{{TOTAL_NPV_ENERGY_CONS}}'),
        (4,  '{{TOTAL_NPV_ENERGY_COST}}'),
        (5,  '{{TOTAL_NPV_CO2_CONS}}'),
        (6,  '{{TOTAL_NPV_SAVINGS_KWH}}'),
        (7,  '{{ANNUAL_SAVINGS}}'),
        (8,  '{{TOTAL_NPV_SAVING_PCT}}'),
        (9,  '{{NPV_INVEST}}'),
        (10, '{{PAYBACK}}'),
        (11, '{{TOTAL_NPV_CO2_AVOIDED}}'),
    ]:
        xml = replace_table_cell(xml, 3, 46, cell_idx, token)

    return xml


# ── ZIP read/write helpers ────────────────────────────────────────────────────

def read_docx(path):
    """Return dict {member_name: bytes} for all entries in the zip."""
    files = {}
    with zipfile.ZipFile(path, 'r') as z:
        for name in z.namelist():
            files[name] = z.read(name)
    return files


def write_docx(files, path):
    """Write the files dict back as a zip (.docx)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    path.write_bytes(buf.getvalue())


# ── Verification ──────────────────────────────────────────────────────────────

EXPECTED_TOKENS = {
    'executive': [
        # Cover
        '{{CUSTOMER}}', '{{PLANT}}', '{{DATA_SOURCE}}',
        '{{REPORT_DATE}}', '{{NPV_COUNT}}', '{{TOTAL_ASSETS}}',
        '{{ELEC_PRICE}}', '{{CURRENCY}}', '{{CO2_INTENSITY}}',
        # Summary
        '{{TOP10_ANNUAL_SAVINGS}}', '{{ANNUAL_SAVINGS}}',
        '{{TOP10_INVEST}}', '{{NPV_INVEST}}',
        '{{TOP10_NPV}}', '{{NPV_VALUE}}',
        '{{TOP10_PAYBACK}}', '{{PAYBACK}}', '{{IRR_DISPLAY}}',
        '{{CO2_SAVINGS}}', '{{BEV_COUNT}}',
        # Table[3] total rows
        '{{TOTAL_TOP10_ENERGY_CONS}}', '{{TOTAL_TOP10_ENERGY_COST}}',
        '{{TOTAL_TOP10_CO2_CONS}}', '{{TOTAL_TOP10_SAVINGS_KWH}}',
        '{{TOTAL_TOP10_SAVING_PCT}}', '{{TOTAL_TOP10_CO2_AVOIDED}}',
        '{{TOTAL_NPV_ENERGY_CONS}}', '{{TOTAL_NPV_ENERGY_COST}}',
        '{{TOTAL_NPV_CO2_CONS}}', '{{TOTAL_NPV_SAVINGS_KWH}}',
        '{{TOTAL_NPV_SAVING_PCT}}', '{{TOTAL_NPV_CO2_AVOIDED}}',
        # Structural
        '{{ROW_TEMPLATE}}', '{{REPORT_DATE_ISO}}',
    ],
    'complete': [
        # Cover
        '{{CUSTOMER}}', '{{PLANT}}', '{{DATA_SOURCE}}',
        '{{REPORT_DATE}}', '{{NPV_COUNT}}', '{{TOTAL_ASSETS}}',
        '{{ELEC_PRICE}}', '{{CURRENCY}}', '{{CO2_INTENSITY}}',
        # Summary
        '{{ANNUAL_SAVINGS}}',
        '{{NPV_INVEST}}', '{{NPV_VALUE}}',
        '{{PAYBACK}}', '{{IRR_DISPLAY}}',
        '{{CO2_SAVINGS}}', '{{BEV_COUNT}}',
        # Table[3] total row
        '{{TOTAL_NPV_ENERGY_CONS}}', '{{TOTAL_NPV_ENERGY_COST}}',
        '{{TOTAL_NPV_CO2_CONS}}', '{{TOTAL_NPV_SAVINGS_KWH}}',
        '{{TOTAL_NPV_SAVING_PCT}}', '{{TOTAL_NPV_CO2_AVOIDED}}',
        # Structural
        '{{ROW_TEMPLATE}}', '{{REPORT_DATE_ISO}}',
    ],
}

SPAIN_VALUES_SHOULD_BE_GONE = {
    'executive': [
        'Global Switch', 'Spain', '06.23.2026',
        # Summary values
        '26,529', '49,928', '178,280', '184,638',
        # TR[11] Top-10 total row
        '5,85,356', '81,950', '87,803', '1,89,495', '32%',
        '1.9', '28,424',
        # TR[12] NPV-positive total row
        '9,61,687', '1,34,636', '1,44,253', '2,23,067',
        '31,229', '23%', '89,528', '2.9', '33,460',
    ],
    'complete': [
        'Global Switch', 'Spain',
        # Summary value
        '89,528',
        # TR[46] NPV-positive total row
        '9,61,687', '1,34,636', '1,44,253', '2,23,067',
        '31,229', '23%', '2.9', '33,460',
    ],
}


def verify(template_key, doc_xml, footer_xml_list):
    full_text = doc_xml + ''.join(footer_xml_list)
    missing  = [t for t in EXPECTED_TOKENS[template_key] if t not in full_text]
    # Check only inside <w:t> nodes to avoid false-positives from CSS/XML attributes
    leftover = [v for v in SPAIN_VALUES_SHOULD_BE_GONE.get(template_key, [])
                if re.search(r'<w:t[^>]*>' + re.escape(v), doc_xml)]
    return missing, leftover


# ── Main ──────────────────────────────────────────────────────────────────────

TOKENIZE = {
    'executive': tokenize_executive,
    'complete':  tokenize_complete,
}

def process(template_key):
    fname   = FILES[template_key]
    src     = TEMPLATES_DIR / fname
    backup  = TEMPLATES_DIR / (fname + '.bak')

    print(f'\n[{template_key}] {fname}')

    # Backup
    if not backup.exists():
        shutil.copy2(src, backup)
        print(f'  Backup → {backup.name}')
    else:
        print(f'  Backup already exists, skipping copy')

    files   = read_docx(src)
    doc_xml = files['word/document.xml'].decode('utf-8')

    print('  Applying tokens to document.xml ...')
    doc_xml = TOKENIZE[template_key](doc_xml)
    files['word/document.xml'] = doc_xml.encode('utf-8')

    # Footer files
    footer_xmls = []
    for name in list(files.keys()):
        if re.match(r'word/footer\d+\.xml', name):
            f_xml = files[name].decode('utf-8')
            f_xml = tokenize_footer(f_xml)
            files[name] = f_xml.encode('utf-8')
            footer_xmls.append(f_xml)
            print(f'  Tokenized {name}')

    write_docx(files, src)
    print(f'  Saved → {src.name}')

    # Verify
    missing, leftover = verify(template_key, doc_xml, footer_xmls)
    if missing:
        print(f'  WARNING — tokens not found in output: {missing}')
    else:
        print(f'  OK — all {len(EXPECTED_TOKENS[template_key])} expected tokens present')
    if leftover:
        print(f'  WARNING — Spain values still in document.xml: {leftover}')
    else:
        print(f'  OK — no tracked Spain values remain in document.xml')


if __name__ == '__main__':
    for key in ['executive', 'complete']:
        process(key)
    print('\nPhase 0 complete.')
