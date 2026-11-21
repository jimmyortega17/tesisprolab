"""Aplica los cambios sobre el Word del usuario (Tesis_..._FINAL_4.docx) sin reconstruirlo:
- saca el blanco de relevancia del capítulo I y lo pone en 4.5.2 (prueba del PMV 1)
- actualiza el lienzo 6x6 (6 ideas por pregunta) y la Tabla 9
- renombra la Tabla 1
- quita "Elaboración propia" (Manual Centrum PUCP, APA 7)
- da formato de portada según el manual
- agrega el llamado a cada tabla y figura
- renumera y regenera las listas y los números de página del índice
"""
import copy, re, io, sys
from docx import Document
from docx.shared import Pt, Emu
from docx.oxml.ns import qn
from PIL import Image
from fb_insert import FDoc, para, run, W
from pages import to_pdf, compute

SRC = 'user_final4.docx'
OUT = '../Tesis_ProLab_Cosecha_Viva_FINAL.docx'
FIG = '../figuras/'

def rep_par(p, old, new):
    t = p.text
    if old not in t: return False
    runs = p.runs
    runs[0].text = t.replace(old, new)
    for r in runs[1:]: r._r.getparent().remove(r._r)
    return True

def find(ps, start, style=None):
    for i, p in enumerate(ps):
        if p.text.strip().startswith(start) and (style is None or p.style.name == style): return i
    raise KeyError(start)

def build(pages=None):
    doc = Document(SRC)
    ps = doc.paragraphs

    # ---------- A. quitar el blanco de relevancia del capítulo I ----------
    i0 = find(ps, 'La Figura 1 presenta el lienzo blanco de relevancia')
    esperado = ['La Figura 1 presenta el lienzo blanco', 'Figura 1', 'Lienzo blanco de relevancia de Cosecha Viva', '', 'Nota. Elaboración propia con la plantilla']
    for k, e in enumerate(esperado):
        assert ps[i0 + k].text.strip().startswith(e), (k, ps[i0 + k].text[:60])
    for k in range(5):
        el = ps[i0 + k]._p; el.getparent().remove(el)

    # ---------- B. Tabla 1: ya no se llama "Lienzo Blanco de Relevancia" ----------
    for p in doc.paragraphs:
        rep_par(p, 'Para delimitar el problema se utilizó el Lienzo Blanco de Relevancia, que ordena los factores según su peso en el problema.', 'Para delimitar el problema se utilizó una matriz de relevancia, que ordena los factores según su peso en el problema.')
        rep_par(p, 'Ambos momentos coinciden con el núcleo crítico del Lienzo Blanco de Relevancia.', 'Ambos momentos coinciden con el núcleo crítico de la matriz de relevancia del problema.')
        if p.text.strip() == 'Lienzo Blanco de Relevancia del Problema': rep_par(p, 'Lienzo Blanco de Relevancia del Problema', 'Matriz de Relevancia de los Factores del Problema')
        if p.text.startswith('Tabla 1. Lienzo Blanco de Relevancia del Problema'):
            p.runs[0].text = p.runs[0].text.replace('Lienzo Blanco de Relevancia del Problema', 'Matriz de Relevancia de los Factores del Problema')

    # ---------- C. lienzo 6x6: seis ideas por pregunta ----------
    ps = doc.paragraphs
    for p in ps:
        rep_par(p, 'y para cada una se plantearon ideas de solución. La Tabla 15 resume las preguntas y las ideas más prometedoras.', 'y para cada una se plantearon seis ideas de solución, de modo que el lienzo reúne 6 preguntas por 6 ideas. La Tabla 15 resume las preguntas y sus seis ideas.')
        rep_par(p, 'La Figura 6 reproduce el lienzo 6x6 con las preguntas generadoras, sus ideas y la idea elegida para cada una.', 'La Figura 6 reproduce el lienzo 6x6: seis preguntas generadoras, seis ideas para cada una y la idea elegida en la última fila.')
    seis = {'¿Cómo elevar': 'Asistencia técnica gratuita; kit de fertirriego en comodato; crédito en insumos; plan de nutrición por parcela; compra grupal de insumos; escuelas de campo entre productores',
            '¿Cómo asegurar': 'Compromiso de compra con prima; venta a cooperativas; subasta de microlotes; contrato con precio piso; alianza con tostadurías de Lima; sello de calidad por lote',
            '¿Cómo financiar': 'Preventa al consumidor; adelanto de la empresa; fondo rotatorio; suscripción anual de café; microcrédito con aliado financiero; adopta una finca (padrinazgo)',
            '¿Cómo dar': 'Videos por etapa; código QR por lote; certificación privada; ficha del productor con foto y nombre; fecha de tueste en la bolsa; transmisión en vivo de la cosecha',
            '¿Cómo acercar': 'Tour en la cosecha; visitas virtuales; encuentro con el productor; cata en la finca; día de cosecha para familias; alojamiento rural con el productor',
            '¿Cómo generar': 'Pago por visitante; venta de subproductos; talleres de catación; venta de plantones; almuerzo gastronómico en la finca; pago por guiar recorridos'}
    for t in doc.tables:
        if t.rows[0].cells[0].text.strip() == 'Pregunta generadora':
            for rw in t.rows:
                for k, v in seis.items():
                    if rw.cells[0].text.startswith(k):
                        par = rw.cells[1].paragraphs[0]; par.runs[0].text = v
                        for r in par.runs[1:]: r._r.getparent().remove(r._r)
    # imagen del 6x6
    ps = doc.paragraphs
    ic = find(ps, 'Lienzo 6x6: preguntas generadoras e ideas')
    imgp = ps[ic + 1]
    blip = imgp._p.find('.//' + qn('a:blip')); rid = blip.get(qn('r:embed'))
    new = open(FIG + 'lienzo_6_lienzo_6x6.png', 'rb').read()
    doc.part.related_parts[rid]._blob = new
    w_, h_ = Image.open(io.BytesIO(new)).size
    ext = imgp._p.find('.//' + qn('wp:extent')); cx = int(ext.get('cx')); cy = int(cx * h_ / w_); ext.set('cy', str(cy))
    for e in imgp._p.iter(qn('a:ext')):
        if e.get('cx'): e.set('cx', str(cx)); e.set('cy', str(cy))

    # ---------- D. blanco de relevancia en 4.5.2 (después de la prueba del PMV 1) ----------
    F = FDoc(doc); F.nf = 98; F.at('b45')
    F.p('Con las respuestas de la prueba del PMV de la primera versión se construyó el lienzo blanco de relevancia, que ordena lo que resultó interesante, las críticas constructivas, las nuevas preguntas y las nuevas ideas (Figura 99). Las nuevas preguntas son las hipótesis H5 a H8 del capítulo VI y las nuevas ideas dieron origen a la segunda versión del PMV.')
    F.figure('Lienzo blanco de relevancia de la prueba del PMV de la primera versión', FIG + 'lienzo_1_blanco_relevancia.png', note=None, width=5.6)
    ps = doc.paragraphs
    anchor = ps[find(ps, '4.5.3 Segunda Versión del PMV', 'Heading 3')]._p
    from docx.text.paragraph import Paragraph
    Paragraph(F.groups['b45'][-1], None).paragraph_format.space_after = Pt(12)
    for el in F.groups['b45']: anchor.addprevious(el)

    # ---------- E. quitar "Elaboración propia" ----------
    def notas():
        for p in doc.paragraphs: yield p
        for t in doc.tables:
            for rw in t.rows:
                for c in rw.cells:
                    for p in c.paragraphs: yield p
    for p in list(notas()):
        t = p.text
        if not t.startswith('Nota.') or 'Elaboración propia' not in t: continue
        rest = t[len('Nota.'):].strip()
        frases = re.split(r'(?<=\.)\s+', rest)
        keep = [f for f in frases if not f.startswith('Elaboración propia')]
        if not keep:
            for r in p.runs: r.text = ''
            p.paragraph_format.line_spacing = 1.0
        else:
            runs = p.runs
            runs[1].text = ' ' + ' '.join(keep)
            for r in runs[2:]: r._r.getparent().remove(r._r)

    # ---------- F. portada según el manual (14 pts arriba, 13 pts abajo) ----------
    ps = doc.paragraphs
    def setp(i, text, size, bold=True):
        p = ps[i]; runs = p.runs; runs[0].text = text
        for r in runs[1:]: r._r.getparent().remove(r._r)
        runs[0].font.size = Pt(size); runs[0].bold = bold
        return p
    setp(0, 'PONTIFICIA UNIVERSIDAD CATÓLICA DEL PERÚ', 14); setp(1, 'ESCUELA DE POSGRADO', 14)
    setp(3, 'Cosecha Viva: Preventa de Café Tecnificado con Trazabilidad para Elevar los Ingresos de los Pequeños Productores de Chanchamayo', 14)
    setp(4, 'TESIS PARA OBTENER EL GRADO DE MAGÍSTER EN ADMINISTRACIÓN ESTRATÉGICA DE EMPRESAS OTORGADO POR LA PONTIFICIA UNIVERSIDAD CATÓLICA DEL PERÚ', 14)
    setp(5, 'PRESENTADA POR', 13)
    for i, n in zip((6, 7, 8, 9), ('Laura Pampa, Diego William', 'Quinto Huatuco, Hernan Huber', 'Aguilar Lazo, Richard Humberto', 'Estrada Cruz, Frank')):
        setp(i, n + ', DNI: [completar]', 13, bold=False)
    setp(10, 'ASESOR', 13)
    setp(11, '[Nombre del asesor], DNI: [completar]', 13, bold=False)
    # ORCID, JURADO y lugar: se insertan después del asesor
    base_line = ps[11]._p
    orcid = copy.deepcopy(base_line); jurh = copy.deepcopy(ps[10]._p)
    jl = [copy.deepcopy(base_line) for _ in range(3)]
    base_line.addnext(orcid)
    def settext(el, text):
        ts = list(el.iter(qn('w:t'))); ts[0].text = text
        for t_ in ts[1:]: t_.text = ''
    settext(orcid, 'ORCID [completar]')
    orcid.addnext(jurh); settext(jurh, 'JURADO')
    prev = jurh
    for k, el in enumerate(jl): prev.addnext(el); settext(el, '[Nombre del jurado %d]' % (k + 1)); prev = el
    ps = doc.paragraphs
    setp(find(ps, 'Lima, noviembre de 2026'), 'Surco, noviembre 2026', 13)
    ps = doc.paragraphs
    fin = find(ps, 'Surco, noviembre 2026')
    cierre = ('ESCUELA DE POSGRADO', 'Cosecha Viva:', 'TESIS PARA OBTENER', 'Estrada Cruz, Frank', 'ORCID', '[Nombre del jurado 3]')
    for p in ps[:fin + 1]:
        for r in p.runs:
            if r.font.size is None and p.text.strip(): r.font.size = Pt(13)
        pf = p.paragraph_format; pf.line_spacing = 1.15; pf.space_before = Pt(0)
        pf.space_after = Pt(12) if p.text.strip().startswith(cierre) or not p.text.strip() else Pt(0)
        if p.text.strip() in ('PRESENTADA POR', 'ASESOR', 'JURADO'): pf.space_after = Pt(2)
    ps[0].paragraph_format.space_before = Pt(24)

    # ---------- G. renumerar tablas y figuras ----------
    ps = doc.paragraphs
    cap = re.compile(r'^(Tabla|Figura) (\d+)$')
    seq = {'Tabla': 0, 'Figura': 0}; mapa = {}
    for p in ps:
        m = cap.match(p.text.strip())
        if m and p.runs and p.runs[0].bold:
            seq[m.group(1)] += 1; mapa[(m.group(1), int(m.group(2)))] = seq[m.group(1)]
    def sub(text):
        t = re.sub(r'Tablas (\d+) y (\d+)', lambda m: 'Tablas %d y %d' % (mapa.get(('Tabla', int(m.group(1))), int(m.group(1))), mapa.get(('Tabla', int(m.group(2))), int(m.group(2)))), text)
        return re.sub(r'\b(Tabla|Figura) (\d+)\b', lambda m: '%s %d' % (m.group(1), mapa.get((m.group(1), int(m.group(2))), int(m.group(2)))), t)
    lista = {p._p for p in ps if re.match(r'^(Tabla|Figura) \d+\. ', p.text)}
    def fix(par):
        if par._p in lista: return
        for r in par.runs:
            t0 = r.text
            if not t0: continue
            t1 = re.sub(r'Tablas (\d+) y (\d+)', lambda m: 'Tablas %d y %d' % (mapa.get(('Tabla', int(m.group(1))), int(m.group(1))), mapa.get(('Tabla', int(m.group(2))), int(m.group(2)))), t0)
            t1 = re.sub(r'\b(Tabla|Figura) (\d+)\b', lambda m: '%s %d' % (m.group(1), mapa.get((m.group(1), int(m.group(2))), int(m.group(2)))), t1)
            if t1 != t0: r.text = t1
    for p in doc.paragraphs: fix(p)
    for t in doc.tables:
        for rw in t.rows:
            for c in rw.cells:
                for p in c.paragraphs: fix(p)

    # ---------- H. llamado a cada tabla y figura (Manual: toda tabla o figura debe tener su llamado) ----------
    ps = doc.paragraphs
    i_ap = next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.strip() == 'Apéndices')
    cuerpo = '\n'.join(p.text for p in ps[:i_ap] if not re.match(r'^(Tabla|Figura) \d+(\.|$)', p.text.strip())) + '\n'.join(c.text for t in doc.tables for rw in t.rows for c in rw.cells)
    nuevos = []
    for k, p in enumerate(ps[:i_ap]):
        m = cap.match(p.text.strip())
        if m and p.runs and p.runs[0].bold:
            n = int(m.group(2))
            if not re.search(r'\b%s %d\b' % (m.group(1), n), cuerpo):
                nuevos.append((p._p, m.group(1), n, ps[k + 1].text.strip()))
    for el, kind, n, title in nuevos:
        t = title[0].lower() + title[1:]
        txt = ('La %s %d presenta %s.' if not t.startswith(('tarjeta',)) else 'La %s %d presenta la %s.') % (kind, n, t)
        if kind == 'Figura': txt = 'La Figura %d presenta %s.' % (n, t)
        if t.startswith('tarjeta'): txt = 'La Tabla %d presenta la %s.' % (n, t)
        newp = para(run(txt, 24), '<w:spacing w:after="0" w:before="0" w:line="480"/><w:ind w:firstLine="720"/><w:jc w:val="left"/>')
        el.addprevious(newp)

    # ---------- I. listas de tablas y figuras ----------
    ps = doc.paragraphs
    caps_t, caps_f = [], []
    for k, p in enumerate(ps):
        m = cap.match(p.text.strip())
        if m and p.runs and p.runs[0].bold:
            (caps_t if m.group(1) == 'Tabla' else caps_f).append((int(m.group(2)), ps[k + 1].text.strip()))
    it = find(ps, 'Lista de Tablas', 'Heading 1'); jf = find(ps, 'Lista de Figuras', 'Heading 1')
    ent_t = [p for p in ps[it + 1:jf] if re.match(r'^Tabla \d+\. ', p.text)]
    ent_f = []
    for p in ps[jf + 1:]:
        if re.match(r'^Figura \d+\. ', p.text): ent_f.append(p)
        else: break
    assert len(ent_t) == len(caps_t) and len(ent_f) == len(caps_f), (len(ent_t), len(caps_t), len(ent_f), len(caps_f))
    def pg(key): return (pages or {}).get(key)
    for p, (n, t) in zip(ent_t, caps_t):
        p.runs[0].text = 'Tabla %d. %s' % (n, t)
        if pg(('Tabla', n)):
            p.runs[1].text = '\t%d' % pg(('Tabla', n))
            for r_ in p.runs[2:]: r_._r.getparent().remove(r_._r)
    for p, (n, t) in zip(ent_f, caps_f):
        p.runs[0].text = 'Figura %d. %s' % (n, t)
        if pg(('Figura', n)):
            p.runs[1].text = '\t%d' % pg(('Figura', n))
            for r_ in p.runs[2:]: r_._r.getparent().remove(r_._r)

    # ---------- J. números de página en el índice (campo TOC): valor en caché ----------
    if pages:
        for p in doc.paragraphs:
            if p.style.name.startswith('toc'):
                title = p.text.split('\t')[0].strip()
                v = pages.get(('h', title))
                if v:
                    ts = [t_ for t_ in p._p.iter(qn('w:t'))]
                    for t_ in reversed(ts):
                        if t_.text and t_.text.strip().isdigit(): t_.text = str(v); break
    doc.save(OUT)
    # datos para calcular páginas
    ps = Document(OUT).paragraphs
    entries = []; j = next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.startswith('Capítulo I:'))
    for p in ps[j:]:
        if p.style.name in ('Heading 1', 'Heading 2', 'Heading 3'): entries.append(('h', p.text.strip()))
    return entries, caps_t, caps_f

if __name__ == '__main__':
    e, ct, cf = build(None)
    pdf = to_pdf(OUT); pages, start, n = compute(pdf, e, ct, cf)
    build(pages)
    pdf = to_pdf(OUT); p2, _, n2 = compute(pdf, e, ct, cf)
    dif = {k: (pages[k], p2[k]) for k in pages if p2.get(k) != pages[k]}
    print('páginas', n2, 'diferencias', len(dif), list(dif.items())[:4])
    if dif: build(p2); pdf = to_pdf(OUT)
