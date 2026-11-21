from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.styles['Normal']; s.font.name = 'Times New Roman'; s.font.size = Pt(12)
        s.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        for sec in self.d.sections:
            sec.left_margin = sec.right_margin = Inches(1); sec.top_margin = sec.bottom_margin = Inches(1)
        for n, sz in (('Heading 1', 14), ('Heading 2', 13), ('Heading 3', 12)):
            h = self.d.styles[n]; h.font.name = 'Times New Roman'; h.font.size = Pt(sz); h.font.bold = True
            h.font.color.rgb = RGBColor(0, 0, 0); h.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        self.nt = 41; self.nf = 5   # la tesis llega a la Tabla 41 y a la Figura 5

    def h(self, text, lvl=2): return self.d.add_heading(text, lvl)
    def p(self, text, bold=False, italic=False, size=None, align=None, indent=True):
        para = self.d.add_paragraph(); r = para.add_run(text); r.bold = bold; r.italic = italic
        if size: r.font.size = Pt(size)
        para.paragraph_format.space_after = Pt(6); para.paragraph_format.line_spacing = 1.15
        if align == 'c': para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return para
    def where(self, text):
        para = self.d.add_paragraph(); r = para.add_run('Dónde insertar: ' + text); r.italic = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        pPr = para._p.get_or_add_pPr(); shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'EAF1FB'); pPr.append(shd)
    def bullets(self, items):
        for t in items:
            para = self.d.add_paragraph(style='List Bullet'); para.add_run(t); para.paragraph_format.space_after = Pt(2)
    def caption(self, kind, title):
        if kind == 'Tabla': self.nt += 1; n = self.nt
        else: self.nf += 1; n = self.nf
        a = self.d.add_paragraph(); a.paragraph_format.keep_with_next = True; a.paragraph_format.space_after = Pt(0)
        a.add_run(f'{kind} {n}').bold = True
        b = self.d.add_paragraph(); b.paragraph_format.keep_with_next = True; b.paragraph_format.space_after = Pt(4)
        b.add_run(title).italic = True
        return n
    def note(self, text):
        para = self.d.add_paragraph(); r = para.add_run('Nota. '); r.italic = True; r.font.size = Pt(10)
        r2 = para.add_run(text); r2.font.size = Pt(10); para.paragraph_format.space_after = Pt(10)
    def table(self, title, header, rows, note=None, widths=None, size=9.5):
        n = self.caption('Tabla', title)
        t = self.d.add_table(rows=1, cols=len(header)); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
        for i, x in enumerate(header):
            c = t.rows[0].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(str(x)); r.bold = True; r.font.size = Pt(size)
            sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:fill'), 'E8EFE0'); c._tc.get_or_add_tcPr().append(sh)
        for row in rows:
            cells = t.add_row().cells
            for i, x in enumerate(row):
                cells[i].text = ''; r = cells[i].paragraphs[0].add_run(str(x)); r.font.size = Pt(size)
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths): row.cells[i].width = Inches(w)
        if note: self.note(note)
        else: self.d.add_paragraph().paragraph_format.space_after = Pt(4)
        return n
    def card(self, title, rows, note=None):
        """Tarjeta de dos columnas (etiqueta, contenido) al estilo Strategyzer."""
        n = self.caption('Tabla', title)
        t = self.d.add_table(rows=0, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
        for lab, txt in rows:
            cells = t.add_row().cells
            cells[0].text = ''; r = cells[0].paragraphs[0].add_run(lab); r.bold = True; r.font.size = Pt(9.5)
            sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:fill'), 'E8EFE0'); cells[0]._tc.get_or_add_tcPr().append(sh)
            cells[1].text = ''; cells[1].paragraphs[0].add_run(txt).font.size = Pt(9.5)
            cells[0].width = Inches(1.7); cells[1].width = Inches(4.8)
        if note: self.note(note)
        else: self.d.add_paragraph().paragraph_format.space_after = Pt(4)
        return n
    def figure(self, title, path, note=None, width=6.0):
        n = self.caption('Figura', title)
        para = self.d.add_paragraph(); para.alignment = WD_ALIGN_PARAGRAPH.CENTER; para.add_run().add_picture(path, width=Inches(width))
        if note: self.note(note)
        return n
    def pb(self): self.d.add_page_break()
    def save(self, path): self.d.save(path)
