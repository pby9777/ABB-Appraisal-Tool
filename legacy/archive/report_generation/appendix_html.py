#!/usr/bin/env python3
"""
appendix_html.py
----------------
Converts the Appendix rich-text editor's HTML (see the "Custom Appendix
Content" section in templates/index.html) into native OOXML fragments that
match the EA report template's own styles, for splicing directly into
document.xml — see splice_appendix_content() in generate_report_standard.py
and generate_report_executive.py.

Only the whitelist of tags the editor itself produces is supported: p, h1-h3,
ul/ol > li, table > tr > td, b/strong, i/em, u, br, img (base64 data URI),
and a caption paragraph (marked data-caption="1"). This is user-authored rich
text, not a trusted template, so unknown tags are unwrapped (their text kept)
rather than raising.

Style/numbering IDs below are not invented — they were confirmed present in
Spain_Global Switch_EA_Report_V1*.docx: Heading1/2/3, Normal, ListParagraph,
Caption and TableGrid all exist as real style IDs in word/styles.xml; numId
28 (bullet, "") and 29 (decimal, "%1.") are existing numbering.xml
definitions already used elsewhere in the document body, reused here rather
than fabricating a new numbering definition.
"""

import re
import base64
from html.parser import HTMLParser

EMU_PER_PX = 9525  # 1 px at 96 DPI, matching what browsers report as naturalWidth/Height
DEFAULT_MAX_WIDTH_EMU = 6480175  # this template's page content width (10205 twips * 635)
BULLET_NUM_ID = 28
NUMBERED_NUM_ID = 29
TABLE_WIDTH_TWIPS = 10205  # this template's page content width in twips


def _xe(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;')
             .replace('>', '&gt;').replace('"', '&quot;'))


class _Node:
    __slots__ = ('tag', 'attrs', 'children')

    def __init__(self, tag, attrs):
        self.tag = tag
        self.attrs = dict(attrs)
        self.children = []  # list[_Node | str]


class _TreeBuilder(HTMLParser):
    _VOID = {'br', 'img'}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = _Node('root', {})
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = _Node(tag, attrs)
        self.stack[-1].children.append(node)
        if tag not in self._VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(_Node(tag, attrs))

    def handle_endtag(self, tag):
        # Pop back to the matching open tag; tolerant of stray/unmatched end
        # tags rather than raising, since this is browser-authored HTML.
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                break

    def handle_data(self, data):
        if data:
            self.stack[-1].children.append(data)


def parse_html(html):
    tb = _TreeBuilder()
    tb.feed(html)
    return tb.root


class _OoxmlBuilder:
    def __init__(self, next_rid, next_media_num, max_width_emu):
        self.next_rid = next_rid
        self.next_media_num = next_media_num
        self.next_doc_pr_id = 900000001  # distinctive range, unlikely to collide with real docPr ids
        self.max_width_emu = max_width_emu
        self.media = {}          # 'word/media/imageN.ext' -> bytes
        self.relationships = []  # [(rId, 'media/imageN.ext'), ...]
        self.parts = []          # ordered list of <w:p>/<w:tbl> xml fragments

    # -- inline run extraction -----------------------------------------------
    def _runs(self, node, fmt):
        """Yield (text_or_'__IMG__', fmt_or_img_node) leaf runs for the
        inline content inside node. fmt = {'b':bool,'i':bool,'u':bool,
        'vert':'superscript'|'subscript'|None}."""
        for child in node.children:
            if isinstance(child, str):
                if child:
                    yield (child, dict(fmt))
                continue
            tag = child.tag
            if tag in ('b', 'strong'):
                yield from self._runs(child, {**fmt, 'b': True})
            elif tag in ('i', 'em'):
                yield from self._runs(child, {**fmt, 'i': True})
            elif tag == 'u':
                yield from self._runs(child, {**fmt, 'u': True})
            elif tag == 'sup':
                yield from self._runs(child, {**fmt, 'vert': 'superscript'})
            elif tag == 'sub':
                yield from self._runs(child, {**fmt, 'vert': 'subscript'})
            elif tag == 'br':
                yield ('\n', dict(fmt))
            elif tag == 'img':
                yield ('__IMG__', child)
            else:
                # Generic wrapper (span/div/font/...) — some browsers apply
                # formatting as an inline style rather than a b/i/u/sup/sub
                # tag; pick that up too rather than silently dropping it.
                yield from self._runs(child, self._merge_style_fmt(child, fmt))

    def _merge_style_fmt(self, node, fmt):
        style = node.attrs.get('style', '')
        if not style:
            return fmt
        fmt = dict(fmt)
        if re.search(r'font-weight\s*:\s*(bold|bolder|[6-9]\d\d)', style):
            fmt['b'] = True
        if re.search(r'font-style\s*:\s*italic', style):
            fmt['i'] = True
        if re.search(r'text-decoration[a-z-]*\s*:\s*[^;]*underline', style):
            fmt['u'] = True
        if re.search(r'vertical-align\s*:\s*super', style):
            fmt['vert'] = 'superscript'
        elif re.search(r'vertical-align\s*:\s*sub', style):
            fmt['vert'] = 'subscript'
        return fmt

    def _run_xml(self, text, fmt):
        if text == '\n':
            return '<w:r><w:br/></w:r>'
        rpr = ''
        if fmt.get('b'):
            rpr += '<w:b/>'
        if fmt.get('i'):
            rpr += '<w:i/>'
        if fmt.get('u'):
            rpr += '<w:u w:val="single"/>'
        if fmt.get('vert'):
            rpr += f'<w:vertAlign w:val="{fmt["vert"]}"/>'
        rpr_xml = f'<w:rPr>{rpr}</w:rPr>' if rpr else ''
        return f'<w:r>{rpr_xml}<w:t xml:space="preserve">{_xe(text)}</w:t></w:r>'

    def _align_of(self, node):
        style = node.attrs.get('style', '')
        m = re.search(r'text-align\s*:\s*(left|right|center|justify)', style)
        return m.group(1) if m else None

    def _para_xml(self, node, pstyle, extra_ppr=''):
        align = self._align_of(node)
        jc = f'<w:jc w:val="{align}"/>' if align else ''
        ppr_body = f'<w:pStyle w:val="{pstyle}"/>{jc}{extra_ppr}' if pstyle else f'{jc}{extra_ppr}'
        ppr = f'<w:pPr>{ppr_body}</w:pPr>' if ppr_body else ''
        runs = []
        for item, meta in self._runs(node, {}):
            if item == '__IMG__':
                runs.append(self._image_run(meta))
            else:
                runs.append(self._run_xml(item, meta))
        if not runs:
            runs = ['<w:r><w:t xml:space="preserve"></w:t></w:r>']
        return f'<w:p>{ppr}{"".join(runs)}</w:p>'

    # -- images ---------------------------------------------------------
    def _image_run(self, img_node):
        src = img_node.attrs.get('src', '')
        m = re.match(r'data:image/(png|jpe?g|gif);base64,(.+)$', src, re.DOTALL)
        if not m:
            return ''
        kind = m.group(1)
        ext = 'png' if kind == 'png' else ('jpg' if kind.startswith('jp') else 'gif')
        try:
            data = base64.b64decode(m.group(2))
        except Exception:
            return ''

        try:
            px_w = int(float(img_node.attrs.get('width') or 0))
            px_h = int(float(img_node.attrs.get('height') or 0))
        except (TypeError, ValueError):
            px_w = px_h = 0
        if px_w <= 0 or px_h <= 0:
            px_w, px_h = 400, 300  # fallback if the browser somehow didn't report natural size

        emu_w, emu_h = px_w * EMU_PER_PX, px_h * EMU_PER_PX
        if emu_w > self.max_width_emu:
            scale = self.max_width_emu / emu_w
            emu_w = self.max_width_emu
            emu_h = max(1, round(emu_h * scale))

        media_name = f'appendix_image{self.next_media_num}.{ext}'
        self.next_media_num += 1
        self.media[f'word/media/{media_name}'] = data

        rid = f'rId{self.next_rid}'
        self.next_rid += 1
        self.relationships.append((rid, f'media/{media_name}'))

        doc_pr_id = self.next_doc_pr_id
        self.next_doc_pr_id += 1
        name = f'AppendixImage{doc_pr_id}'
        return (
            '<w:r><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{emu_w}" cy="{emu_h}"/>'
            f'<wp:docPr id="{doc_pr_id}" name="{name}"/>'
            '<wp:cNvGraphicFramePr>'
            '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
            '</wp:cNvGraphicFramePr>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{doc_pr_id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu_w}" cy="{emu_h}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
        )

    # -- lists / tables ---------------------------------------------------
    def _list_xml(self, node, num_id):
        out = []
        for li in node.children:
            if isinstance(li, str) or li.tag != 'li':
                continue
            numpr = f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'
            out.append(self._para_xml(li, 'ListParagraph', numpr))
        return ''.join(out)

    def _table_xml(self, node):
        rows = [tr for tr in node.children if not isinstance(tr, str) and tr.tag == 'tr']
        ncols = max(
            (sum(1 for c in tr.children if not isinstance(c, str) and c.tag in ('td', 'th'))
             for tr in rows),
            default=0,
        )
        if ncols == 0:
            return ''
        col_w = TABLE_WIDTH_TWIPS // ncols
        rows_xml = []
        for tr in rows:
            cells_xml = []
            for td in tr.children:
                if isinstance(td, str) or td.tag not in ('td', 'th'):
                    continue
                para = self._para_xml(td, 'Normal') if td.children else '<w:p/>'
                cells_xml.append(
                    f'<w:tc><w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/></w:tcPr>{para}</w:tc>'
                )
            rows_xml.append(f'<w:tr>{"".join(cells_xml)}</w:tr>')
        grid = ''.join(f'<w:gridCol w:w="{col_w}"/>' for _ in range(ncols))
        return (
            '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
            f'<w:tblW w:w="{TABLE_WIDTH_TWIPS}" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr>'
            f'<w:tblGrid>{grid}</w:tblGrid>{"".join(rows_xml)}</w:tbl>'
            '<w:p/>'  # guards against two adjacent user tables silently merging in Word
        )

    # -- top-level dispatch -------------------------------------------------
    def walk(self, root):
        for node in root.children:
            if isinstance(node, str):
                if node.strip():
                    self.parts.append(
                        f'<w:p><w:r><w:t xml:space="preserve">{_xe(node)}</w:t></w:r></w:p>'
                    )
                continue
            tag = node.tag
            if tag == 'h1':
                self.parts.append(self._para_xml(node, 'Heading1'))
            elif tag == 'h2':
                self.parts.append(self._para_xml(node, 'Heading2'))
            elif tag == 'h3':
                self.parts.append(self._para_xml(node, 'Heading3'))
            elif tag == 'p':
                pstyle = 'Caption' if node.attrs.get('data-caption') == '1' else 'Normal'
                self.parts.append(self._para_xml(node, pstyle))
            elif tag == 'ul':
                self.parts.append(self._list_xml(node, BULLET_NUM_ID))
            elif tag == 'ol':
                self.parts.append(self._list_xml(node, NUMBERED_NUM_ID))
            elif tag == 'table':
                self.parts.append(self._table_xml(node))
            elif tag in ('div', 'section', 'body', 'html'):
                self.walk(node)  # unwrap generic containers
            else:
                self.parts.append(self._para_xml(node, 'Normal'))


def html_to_ooxml(html, next_rid, next_media_num, max_width_emu=DEFAULT_MAX_WIDTH_EMU):
    """Convert Appendix editor HTML into (ooxml_fragment, media_files, relationships).

    next_rid: first unused numeric relationship id in word/_rels/document.xml.rels
              (caller scans for the current max and passes max+1).
    next_media_num: first unused numeric suffix for word/media/appendix_imageN.*
                    (caller scans existing media filenames similarly).
    """
    if not html or not html.strip():
        return '', {}, []
    root = parse_html(html)
    builder = _OoxmlBuilder(next_rid, next_media_num, max_width_emu)
    builder.walk(root)
    return ''.join(builder.parts), builder.media, builder.relationships


# ---------------------------------------------------------------------------
# Splicing into a report's document.xml / rels / [Content_Types].xml
# ---------------------------------------------------------------------------

def _next_rid(rels_xml):
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return (max(ids) + 1) if ids else 1


def _next_media_num(existing_media_names):
    nums = [int(m.group(1)) for name in existing_media_names
            for m in [re.search(r'image(\d+)\.', name)] if m]
    return (max(nums) + 1) if nums else 1


def find_appendix_insertion_point(doc_xml):
    """Char offset right after the real (non-TOC) 'Appendix' Heading1
    paragraph, or None if not found. Requiring pStyle="Heading1" on the
    match is what skips the TOC's own '>Appendix<' entry earlier in the
    document (a TOC1-styled hyperlink, not the real section heading)."""
    for m in re.finditer(r'<w:p [^>]*>(?:(?!</w:p>).)*?</w:p>', doc_xml, re.DOTALL):
        para = m.group()
        if 'pStyle w:val="Heading1"' in para and '>Appendix<' in para:
            return m.end()
    return None


def find_document_end_insertion_point(doc_xml):
    """Char offset right before the document body's final <w:sectPr> — the
    last thing content can be inserted before, since sectPr must remain the
    body's last child. This template's Appendix is the report's last
    section, so this is also, in effect, the end of the Appendix."""
    idx = doc_xml.rfind('<w:sectPr')
    if idx != -1:
        return idx
    idx = doc_xml.rfind('</w:body>')
    return idx if idx != -1 else len(doc_xml)


def _find_section_heading_start(doc_xml, anchor_text):
    """Char offset of the START of the first non-TOC paragraph whose text
    contains anchor_text — the same heading paragraph locate_table() finds
    the following table from, but returning where IT begins rather than
    where its table ends."""
    from report_table_utils import _norm_dash, _is_toc_para, _para_text
    target = _norm_dash(anchor_text).lower()
    for m in re.finditer(r'<w:p[ >].*?</w:p>', doc_xml, re.DOTALL):
        if _is_toc_para(m.group()):
            continue
        if target in _norm_dash(_para_text(m.group())).lower():
            return m.start()
    return None


# Keys the UI's "insert custom content" dropdown may send; anything else
# (including omitted/unrecognized values) falls back to 'appendix_start' —
# this feature's original, only, behavior — so this is not a breaking change
# for any older caller.
INSERTION_POSITIONS = ('appendix_start', 'appendix_end', 'before_details', 'after_details')


def resolve_insertion_point(doc_xml, position, details_anchor):
    """Char offset to splice custom Appendix content at, for one of
    INSERTION_POSITIONS. details_anchor is the same anchor text the caller
    already passes to report_table_utils.render_table_section() for its
    (Standard's one, or Executive's Top-10) Application Details table."""
    from report_table_utils import locate_table

    if position == 'appendix_end':
        return find_document_end_insertion_point(doc_xml)

    if position == 'before_details' and details_anchor:
        pt = _find_section_heading_start(doc_xml, details_anchor)
        if pt is not None:
            return pt

    if position == 'after_details' and details_anchor:
        _, te = locate_table(doc_xml, details_anchor)
        if te is not None:
            return te

    return find_appendix_insertion_point(doc_xml)


def ensure_image_content_types(content_types_xml):
    """Add Default Extension entries for jpg/jpeg/gif if a template doesn't
    already carry them — png is already registered in every template this
    codebase ships, but user-supplied images could be any of these formats
    (the editor prefers PNG, but this guards against an OPC-invalid package
    — an unregistered part extension is a real corruption Word will reject —
    regardless of what the frontend actually sends)."""
    needed = {
        'jpg':  'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif':  'image/gif',
    }
    for ext, content_type in needed.items():
        if f'Extension="{ext}"' not in content_types_xml:
            entry = f'<Default Extension="{ext}" ContentType="{content_type}"/>'
            content_types_xml = content_types_xml.replace('</Types>', entry + '</Types>')
    return content_types_xml


def splice_appendix(doc_xml, rels_xml, appendix_html_path, existing_media_names,
                     position='appendix_start', details_anchor=None):
    """Insert the Appendix editor's custom content at one of
    INSERTION_POSITIONS (default: right after the real Appendix heading,
    this feature's original behavior).

    Returns (doc_xml, rels_xml, media_files) — media_files is {} and
    doc_xml/rels_xml come back unchanged if appendix_html_path is falsy,
    the file is empty, or the target insertion point can't be located.
    """
    if not appendix_html_path:
        return doc_xml, rels_xml, {}

    with open(appendix_html_path, 'r', encoding='utf-8') as f:
        html = f.read()
    if not html.strip():
        return doc_xml, rels_xml, {}

    insert_at = resolve_insertion_point(doc_xml, position, details_anchor)
    if insert_at is None:
        print("  WARNING: Appendix insertion point not found — custom Appendix content skipped.")
        return doc_xml, rels_xml, {}

    ooxml, media, new_rels = html_to_ooxml(
        html,
        next_rid=_next_rid(rels_xml),
        next_media_num=_next_media_num(existing_media_names),
    )
    if not ooxml:
        return doc_xml, rels_xml, {}

    doc_xml = doc_xml[:insert_at] + ooxml + doc_xml[insert_at:]

    rel_entries = ''.join(
        f'<Relationship Id="{rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="{target}"/>'
        for rid, target in new_rels
    )
    rels_xml = rels_xml.replace('</Relationships>', rel_entries + '</Relationships>')

    print(f"  Custom Appendix content: {len(media)} image(s), inserted at position '{position}'.")
    return doc_xml, rels_xml, media
