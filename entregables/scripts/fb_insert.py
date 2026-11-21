import copy, re
from lxml import etree
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import parse_xml

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FONT = '<w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman"/>'
def esc(t): return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
def run(text, sz=24, b=False, i=False):
    rp = FONT + ('<w:b/><w:bCs/>' if b else '') + ('<w:i/><w:iCs/>' if i else '') + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    return f'<w:r><w:rPr>{rp}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
def para(runs, ppr=''): return parse_xml(f'<w:p {W}><w:pPr>{ppr}</w:pPr>{runs}</w:p>')

class FDoc:
    """Adaptador con la misma API que Doc, pero escribe con el formato de la tesis original."""
    def __init__(self, doc):
        self.doc = doc; self.body = doc.element.body; self.nt = 41; self.nf = 5
        self.groups = {}; self.key = None; self.created = []
    # --- control de grupos
    def at(self, key): self.key = key; self.groups.setdefault(key, [])
    def _add(self, el):
        sect = self.body[-1]  # sectPr
        sect.addprevious(el)
        if self.key: self.groups[self.key].append(el)
        return el
    # --- API de contenido
    def h(self, text, lvl=2):
        if lvl == 1 and not text.startswith('Apéndice'): return
        if lvl == 1: lvl = 2
        style = {2: 'Heading2', 3: 'Heading3'}[lvl]
        b_i = '<w:b/><w:bCs/>' + ('<w:i/><w:iCs/>' if lvl == 3 else '')
        self._add(parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:rPr>{FONT}{b_i}<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'))
    def p(self, text, bold=False, italic=False, size=None, align=None, indent=True):
        if text.startswith(('Referencia nueva para la lista', 'Este documento reúne')): return
        sz = 24 if not size else 24
        self._add(para(run(text, sz, bold, italic), '<w:spacing w:after="0" w:before="0" w:line="480"/>' + ('<w:ind w:firstLine="720"/>' if indent else '') + '<w:jc w:val="left"/>'))
    def where(self, text): pass
    def pb(self): pass
    def bullets(self, items):
        for t in items:
            self._add(para(run('•\t' + t), '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:spacing w:after="0" w:before="0" w:line="480"/><w:ind w:left="720" w:hanging="360"/>'))
    def caption(self, kind, title):
        if kind == 'Tabla': self.nt += 1; n = self.nt
        else: self.nf += 1; n = self.nf
        self._add(para(run(f'{kind} {n}', 24, b=True), '<w:keepNext/><w:spacing w:after="0" w:before="240" w:line="480"/>'))
        self._add(para(run(title, 24, i=True), '<w:keepNext/><w:spacing w:after="0" w:before="0" w:line="480"/>'))
        return n
    def note(self, text):
        self._add(para(run('Nota.', 24, i=True) + run(' ' + text), '<w:spacing w:after="240" w:before="0" w:line="480"/>'))
    def _tbl(self, header, rows, widths, size, first_col_bold=False):
        tot = 9026; ws = widths or [1] * len(header); s = sum(ws); dx = [int(tot * w / s) for w in ws]
        sz = int(size * 2)
        def tc(text, w, top, bot, hdr=False, keep=True):
            borders = f'<w:tcBorders><w:top w:val="{"single" if top else "none"}" w:color="000000" w:sz="6"/><w:left w:val="none" w:color="FFFFFF" w:sz="0"/><w:bottom w:val="{"single" if bot else "none"}" w:color="000000" w:sz="6"/><w:right w:val="none" w:color="FFFFFF" w:sz="0"/></w:tcBorders>'
            mar = '<w:tcMar><w:top w:type="dxa" w:w="40"/><w:left w:type="dxa" w:w="80"/><w:bottom w:type="dxa" w:w="40"/><w:right w:type="dxa" w:w="80"/></w:tcMar>'
            return (f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{w}"/>{borders}{mar}</w:tcPr>'
                    f'<w:p><w:pPr>{"<w:keepNext/>" if keep else ""}<w:spacing w:after="20" w:before="20" w:line="240"/><w:jc w:val="left"/></w:pPr>{run(str(text), sz, b=hdr)}</w:p></w:tc>')
        x = f'<w:tbl {W}><w:tblPr><w:tblW w:type="dxa" w:w="{tot}"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="none" w:color="FFFFFF" w:sz="0"/><w:left w:val="none" w:color="FFFFFF" w:sz="0"/><w:bottom w:val="none" w:color="FFFFFF" w:sz="0"/><w:right w:val="none" w:color="FFFFFF" w:sz="0"/><w:insideH w:val="none" w:color="FFFFFF" w:sz="0"/><w:insideV w:val="none" w:color="FFFFFF" w:sz="0"/></w:tblBorders></w:tblPr><w:tblGrid>' + ''.join(f'<w:gridCol w:w="{d}"/>' for d in dx) + '</w:tblGrid>'
        if header:
            x += '<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>' + ''.join(tc(h, d, True, True, hdr=True) for h, d in zip(header, dx)) + '</w:tr>'
        for k, r in enumerate(rows):
            last = k == len(rows) - 1
            x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + ''.join(tc(c, d, False, last, hdr=(first_col_bold and j == 0), keep=(k == 0)) for j, (c, d) in enumerate(zip(r, dx))) + '</w:tr>'
        return parse_xml(x + '</w:tbl>')
    def table(self, title, header, rows, note=None, widths=None, size=9.5):
        n = self.caption_if('Tabla', title)
        if n is None: return
        self._add(self._tbl(header, rows, widths, min(size, 9)))
        self.note(note) if note else None
    def caption_if(self, kind, title):
        if title.startswith('Mapa de inserción'): self.nt += 1; return None
        return self.caption(kind, title)
    def card(self, title, rows, note=None):
        self.caption('Tabla', title)
        self._add(self._tbl(None, rows, [1.7, 4.8], 9, first_col_bold=True))
        self.note(note) if note else None
    def figure(self, title, path, note=None, width=6.0):
        self.caption('Figura', title)
        p = self.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path, width=Inches(min(width, 5.6)))
        el = p._p; el.getparent().remove(el); self._add(el)
        self.note(note) if note else None

def move_groups(F, anchors):
    """anchors: {clave: elemento ancla}; mueve cada grupo antes de su ancla (o al final si el ancla es None)."""
    for key, els in F.groups.items():
        anc = anchors.get(key)
        if anc is None: continue
        for el in els: anc.addprevious(el)
    # elementos sin clave o grupos sin ancla se eliminan
    used = set(id(e) for k, es in F.groups.items() if anchors.get(k) is not None for e in es)
    for k, es in F.groups.items():
        if anchors.get(k) is None:
            for e in es: e.getparent().remove(e)
