"""Reestructura el Word de Cosecha Viva con la estructura de títulos de ProLab (modelo: tesis AnsioZen).
Parte de user15.docx. No crea títulos que no existan en el modelo; el contenido de los títulos eliminados se
reubica dentro de la sección que corresponde o en un apéndice.
"""
import copy, re, sys
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.text.paragraph import Paragraph
from fb_insert import FDoc, para, run, W
from fb_data import VS, BS, n0, row
from pages import to_pdf, compute

SRC = 'user15.docx'
OUT = '../Tesis_ProLab_Cosecha_Viva_FINAL.docx'
FIG = '../figuras/'

def build(pages=None):
    doc = Document(SRC)
    body = doc.element.body
    B = list(body.iterchildren())
    P = lambda i: Paragraph(B[i], doc)
    def chk(i, start):
        t = P(i).text.strip() if B[i].tag.endswith('}p') else ''
        assert t.startswith(start), (i, t[:60], start)
    for i, s in {205: 'Capítulo I:', 213: 'Para delimitar', 218: 'La Figura 1 delimita', 225: '1.3.1', 246: '2.2', 252: 'El análisis de las cinco', 254: '2.3', 260: '2.4', 261: 'La sección 2.2', 266: '2.5',
                 273: 'Capítulo III', 299: '3.1.1', 311: '3.2', 323: '3.2.1', 345: 'La Figura 5', 355: 'La Figura 6', 379: 'La Figura 7', 385: '4.4.1', 397: '4.5', 428: '5.3.1',
                 437: '5.5', 474: '5.6', 480: 'Capítulo VI', 489: '6.1.1', 496: '6.1.2', 512: '6.1.4', 547: '6.1.5', 553: '6.2', 581: '6.3.3', 619: '6.3.8', 637: '6.3.10', 676: '6.3.11',
                 696: '7.2', 703: 'El VAN social se calculó', 708: 'El VAN social asciende', 709: '7.3', 715: '7.4', 738: '8.1.1', 850: 'Apéndice H'}.items(): chk(i, s)

    def clone_heading(like, text):
        el = copy.deepcopy(B[like])
        for bm in list(el.iter(qn('w:bookmarkStart'))) + list(el.iter(qn('w:bookmarkEnd'))): bm.getparent().remove(bm)
        ts = list(el.iter(qn('w:t'))); ts[0].text = text
        for t_ in ts[1:]: t_.text = ''
        return el
    def clone_par(like, text):
        el = copy.deepcopy(B[like])
        for bm in list(el.iter(qn('w:bookmarkStart'))) + list(el.iter(qn('w:bookmarkEnd'))): bm.getparent().remove(bm)
        rs = list(el.iter(qn('w:r')))
        for r in rs[1:]: r.getparent().remove(r)
        ts = list(rs[0].iter(qn('w:t'))); ts[0].text = text
        for t_ in ts[1:]: t_.text = ''
        return el
    H2, H3, TXT = 224, 225, 207
    def h3(text): return clone_heading(H3, text)
    def h2(text): return clone_heading(850, text)
    def txt(text): return clone_par(TXT, text)
    def before(anchor, els):
        for e in els: anchor.addprevious(e)
    def rm(els):
        for e in els:
            if e.getparent() is not None: e.getparent().remove(e)

    # ================= Capítulo I: 1.2.1 Lienzo de Dos Dimensiones =================
    B[213].addprevious(h3('1.2.1 Lienzo de Dos Dimensiones: Delimitación del Problema'))
    p1 = clone_par(218, 'Para precisar el problema se aplica el Lienzo de Dos Dimensiones, que distingue lo que forma parte del problema (qué es) y lo que queda fuera de su alcance (qué no es). La Figura 1 presenta el lienzo.')
    p2 = txt('Qué es: la baja productividad (cerca de 15 qq/ha frente a un potencial de 50), la venta de café sin diferenciar al intermediario a un precio bajo y volátil, la falta de capital y de un comprador seguro, y la migración de jóvenes y el abandono de fincas.')
    p3 = txt('Qué no es: no es solo un problema de precio, porque sin tecnificación no hay calidad que pagar; no es solo un problema de productividad, porque sin comprador no hay incentivo a invertir; no es un problema que resuelva el turismo cafetalero por sí solo; y no depende del Estado, porque se resuelve con recursos de mercado.')
    B[218].addprevious(p1); B[218].addprevious(p2); B[218].addprevious(p3); rm([B[218]])

    # ================= Capítulo II =================
    seg = B[267:273]; pestel = B[255:260]; legal = B[475:480]
    mk_head = B[437]; mk = B[438:474]
    before(B[246], seg)                      # segmentación dentro de 2.1
    rm([B[266], B[254], B[260], B[265]])
    # 2.2: Porter (2.2.1) y competidores (2.2.2)
    por_intro = clone_par(261, 'Las cinco fuerzas de Porter se evaluaron para el negocio de café de especialidad con preventa en Lima. La Tabla 6 resume la intensidad de cada fuerza y la respuesta del modelo. La valoración corresponde al juicio del equipo y debe contrastarse con entrevistas a tostadurías y cafeterías.')
    rm([B[261]])
    intro22 = txt('El análisis competitivo se realiza con las cinco fuerzas de Porter y con la comparación de competidores directos e indirectos.')
    seq = [intro22, h3('2.2.1 Fuerzas Competitivas: Análisis de Cinco Fuerzas de Porter'), por_intro, B[262], B[263], B[264], B[252],
           h3('2.2.2 Mapa de Posicionamiento Estratégico y Análisis Competitivo'), B[247], B[248], B[249], B[250], B[251], B[253]]
    before(B[273], seq)

    # ================= Capítulo III =================
    meta_fig = B[282:287]; empat = B[300:311]
    rm([B[299]])
    B[311].addprevious(txt('El lienzo de meta-usuario de Rosa y los mapas de empatía del productor y del consumidor se presentan en el Apéndice J.'))
    rm([B[323]])

    # ================= Capítulo IV =================
    fig5 = B[345:350]; fig6 = B[355:360]; fig7 = B[379:384]; vpc = B[386:397]
    rm([B[385]])
    B[350].addprevious(txt('El lienzo 6x6 se presenta en el Apéndice J.'))
    B[360].addprevious(txt('La matriz costo-impacto se presenta en el Apéndice J.'))
    B[397].addprevious(txt('El lienzo de la propuesta de valor del productor y del consumidor se presenta en el Apéndice J.'))

    # ================= Capítulo V =================
    rm([B[428], B[437], B[474]])
    def runin(like, text):
        el = clone_par(like, text)
        for r in el.iter(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is None: rpr = parse_xml('<w:rPr %s/>' % W.replace('xmlns:w=', 'xmlns:w=')); r.insert(0, rpr)
            if rpr.find(qn('w:b')) is None: rpr.insert(0, parse_xml('<w:b %s/>' % W))
        return el
    mk_new = []
    for i in range(438, 474):
        e = B[i]
        if e.tag.endswith('}p') and P(i).style.name == 'Heading 3':
            label = re.sub(r'^5\.5\.\d+\s+', '', P(i).text.strip()) + '.'
            mk_new.append(runin(TXT, label)); rm([e])
        else:
            mk_new.append(e)

    # ================= Capítulo VI =================
    prior = B[497:506]; diseno = B[507:512]; tareas = B[548:553]; tarjetas = B[513:547]
    rm([B[496], B[506], B[512], B[547]])
    call_k = txt('Las tarjetas de prueba y de aprendizaje de las hipótesis H1 a H8 se presentan en el Apéndice K.')
    before(B[489], prior + diseno + [call_k] + tareas)
    # 6.2: Plan de Mercadeo y Plan de Operaciones
    intro62 = txt('La factibilidad se evalúa con el plan de mercadeo y el plan de operaciones.')
    before(B[554], [intro62, h3('6.2.1 Plan de Mercadeo')] + mk_new + [h3('6.2.2 Plan de Operaciones')])
    # 6.3
    rm([B[637], B[676]])
    ind = B[677:683]; fin = B[683:688]
    before(B[619], ind)
    before(B[581], fin)

    # ================= Capítulo VII =================
    amb = B[710:715]; gen = B[716:721]
    rm([B[709], B[715]])
    before(B[696], amb + gen)
    F = FDoc(doc); F.nt = 99; F.nf = 98
    F.at('ben')
    F.p('Los beneficios sociales son el ingreso bruto adicional de los productores frente a la situación sin el modelo, valorado a precios de mercado. Cada ODS declarado genera un beneficio: el ODS 2 (meta 2.3) genera el valor de la producción adicional que logra el productor al pasar de 15 a 25 quintales por hectárea (20 quintales por productor con tecnificación madura, al precio del intermediario de S/ 690 por quintal); el ODS 1 (meta 1.2) genera el sobreprecio que recibe por su café selecto (S/ 230 por quintal frente al intermediario); y el ODS 8 (metas 8.3 y 8.9) genera la remuneración del empleo calificado de los técnicos de campo y el gasto local de los tours (transporte, guía, almuerzo y pago al productor). La Tabla 100 presenta la proyección.')
    hdr = ['Concepto', 'Año 1', 'Año 2', 'Año 3', 'Año 4', 'Año 5']
    ben = [['Beneficio social 1 (ODS 2): valor de la producción adicional'] + [n0(v) for v in row(BS, 11, 2, 6)],
           ['Beneficio social 2 (ODS 1): sobreprecio del café selecto'] + [n0(v) for v in row(BS, 16, 2, 6)],
           ['Beneficio social 3 (ODS 8): empleo calificado de campo'] + [n0(v) for v in row(BS, 20, 2, 6)],
           ['Beneficio social 4 (ODS 8): gasto local de los tours'] + [n0(v) for v in row(BS, 22, 2, 6)],
           ['Total de beneficios sociales'] + [n0(v) for v in row(BS, 23, 2, 6)]]
    F.table('Beneficios Sociales por ODS (S/)', hdr, ben, note='Adaptado del modelo financiero (hoja Beneficios Sociales).', widths=[3.0, 0.9, 0.9, 0.9, 0.9, 0.9])
    ben_els = F.groups['ben']
    F.at('cos')
    F.p('El VAN social se calcula como la diferencia entre los beneficios sociales y los costos sociales, descontada a la tasa social del MEF de 8%. No incluye el flujo de caja libre ni la inversión de la empresa, porque mide el valor que se crea para la sociedad y no el retorno del inversionista. Los costos sociales son los recursos adicionales que el productor destina a insumos y cosecha selectiva (S/ 5,000 y S/ 2,500 por año). No se aplicaron precios sociales a la mano de obra ni a las divisas. La Tabla 101 presenta el cálculo.')
    cos = [['Costos sociales'] + [n0(v) for v in row(VS, 6, 2, 6)],
           ['Flujo social (beneficios - costos)'] + [n0(v) for v in row(VS, 7, 2, 6)],
           ['Valor presente al 8%'] + [n0(VS.cell(11 + i, 2).value) for i in range(5)]]
    F.table('Costos Sociales, Flujo Social y VAN Social al 8% (S/)', hdr, cos, note='Adaptado del modelo financiero (hoja VAN Social). VAN social = suma de los valores presentes del flujo social; no incluye flujo de caja libre ni inversión.', widths=[3.0, 0.9, 0.9, 0.9, 0.9, 0.9])
    cos_els = F.groups['cos']
    rm(B[703:708])
    before(B[698], [h3('7.2.1 Beneficios Sociales')])
    before(B[708], ben_els + [h3('7.2.2 Costos Sociales')] + cos_els)

    # ================= Capítulo VIII =================
    rm([B[738]])

    # ================= Apéndices nuevos =================
    sect = body[-1]
    def add(els):
        for e in els: sect.addprevious(e)
    add([h2('Apéndice I. Análisis PESTEL Completo del Mercado Peruano'), clone_par(851, 'Este apéndice presenta el análisis del entorno político, económico, social, tecnológico, ecológico y legal en el que opera Cosecha Viva, y el marco legal y regulatorio del modelo.')] + pestel + legal)
    add([h2('Apéndice J. Lienzos: Meta Usuario, Propuesta de Valor, Lienzo 6x6 y Costo vs Impacto'), clone_par(851, 'Este apéndice reúne los lienzos elaborados durante el proceso de tesis.'),
         h3('J.1 Lienzo de Meta Usuario')] + meta_fig + empat + [h3('J.2 Lienzo de Propuesta de Valor')] + fig7 + vpc + [h3('J.3 Lienzo 6x6')] + fig5 + [h3('J.4 Lienzo Costo vs Impacto')] + fig6)
    add([h2('Apéndice K. Tarjetas de Prueba y Aprendizaje'), clone_par(851, 'Este apéndice presenta las tarjetas de prueba y de aprendizaje de las hipótesis H1 a H8.')] + tarjetas)

    # ================= correcciones de texto =================
    def rep(par, old, new):
        t = par.text
        if old not in t: return False
        rs = par.runs; rs[0].text = t.replace(old, new)
        for r in rs[1:]: r._r.getparent().remove(r._r)
        return True
    def all_pars():
        for p in doc.paragraphs: yield p
        for t in doc.tables:
            for rw in t.rows:
                for c in rw.cells:
                    for p in c.paragraphs: yield p
    for p in all_pars():
        rep(p, '(sección 2.5)', '(sección 2.1)'); rep(p, '(sección 5.6)', '(Apéndice I)')
        rep(p, ' Guion completo en el Apéndice J.', '')
        rep(p, 'definidos antes de su aplicación. La Tabla', 'definidos antes de su aplicación. La Tabla') if False else None
        if p.text.strip() == 'Nota.':
            for r in p.runs: r.text = ''
            p.paragraph_format.line_spacing = 1.0

    # ================= renumerar tablas y figuras (orden del documento) =================
    cap = re.compile(r'^(Tabla|Figura) (\d+)$')
    ps = doc.paragraphs; seq_ = {'Tabla': 0, 'Figura': 0}; mapa = {}
    for p in ps:
        m = cap.match(p.text.strip())
        if m and p.runs and p.runs[0].bold:
            seq_[m.group(1)] += 1; mapa[(m.group(1), int(m.group(2)))] = seq_[m.group(1)]
    lista = {p._p for p in ps if re.match(r'^(Tabla|Figura) \d+\. ', p.text)}
    def fix(par):
        if par._p in lista: return
        for r in par.runs:
            t0 = r.text
            if not t0: continue
            t1 = re.sub(r'Tablas (\d+) y (\d+)', lambda m: 'Tablas %d y %d' % (mapa.get(('Tabla', int(m.group(1))), int(m.group(1))), mapa.get(('Tabla', int(m.group(2))), int(m.group(2)))), t0)
            t1 = re.sub(r'\b(Tabla|Figura) (\d+)\b', lambda m: '%s %d' % (m.group(1), mapa.get((m.group(1), int(m.group(2))), int(m.group(2)))), t1)
            if t1 != t0: r.text = t1
    for p in all_pars(): fix(p)

    # ================= llamado a cada tabla y figura (cuerpo) =================
    ps = doc.paragraphs
    i_ap = next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.strip() == 'Apéndices')
    cuerpo = '\n'.join(p.text for p in ps if not re.match(r'^(Tabla|Figura) \d+(\.|$)', p.text.strip())) + '\n'.join(c.text for t in doc.tables for rw in t.rows for c in rw.cells)
    nuevos = []
    for k, p in enumerate(ps):
        m = cap.match(p.text.strip())
        if m and p.runs and p.runs[0].bold:
            n = int(m.group(2))
            if not re.search(r'\b%s %d\b' % (m.group(1), n), cuerpo): nuevos.append((p._p, m.group(1), n, ps[k + 1].text.strip(), k < i_ap))
    for el, kind, n, title, enbody in nuevos:
        if not enbody: continue
        t = title[0].lower() + title[1:]
        txt_ = ('La Tabla %d presenta la %s.' if t.startswith('tarjeta') else '%s %d presenta %s.') % ((n, t) if t.startswith('tarjeta') else (('La ' + kind), n, t))
        el.addprevious(para(run(txt_, 24), '<w:spacing w:after="0" w:before="0" w:line="480"/><w:ind w:firstLine="720"/><w:jc w:val="left"/>'))

    # ================= listas de tablas y figuras =================
    ps = doc.paragraphs
    caps_t, caps_f = [], []
    for k, p in enumerate(ps):
        m = cap.match(p.text.strip())
        if m and p.runs and p.runs[0].bold:
            (caps_t if m.group(1) == 'Tabla' else caps_f).append((int(m.group(2)), ps[k + 1].text.strip()))
    def idx_h(text):
        return next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.strip() == text)
    it, jf = idx_h('Lista de Tablas'), idx_h('Lista de Figuras')
    ent_t = [p for p in ps[it + 1:jf] if re.match(r'^Tabla \d+\. ', p.text)]
    ent_f = []
    for p in ps[jf + 1:]:
        if re.match(r'^Figura \d+\. ', p.text): ent_f.append(p)
        else: break
    while len(ent_t) < len(caps_t):
        nuevo = copy.deepcopy(ent_t[-1]._p); ent_t[-1]._p.addnext(nuevo); ent_t.append(Paragraph(nuevo, doc))
    while len(ent_f) < len(caps_f):
        nuevo = copy.deepcopy(ent_f[-1]._p); ent_f[-1]._p.addnext(nuevo); ent_f.append(Paragraph(nuevo, doc))
    assert len(ent_t) == len(caps_t) and len(ent_f) == len(caps_f)
    pg = lambda k: (pages or {}).get(k)
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

    # ================= índice (campo TOC): se regenera con los títulos actuales =================
    ps = doc.paragraphs
    toc_old = [p for p in ps if p.style.name.startswith('toc')]
    roman = {}
    for p in toc_old:
        tt, _, pgv = p.text.partition('\t')
        if pgv and not pgv.strip().isdigit(): roman[tt.strip()] = pgv.strip()
    heads = []
    start_h = next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.strip() == 'Lista de Tablas')
    for p in ps[start_h:]:
        if p.style.name in ('Heading 1', 'Heading 2', 'Heading 3'): heads.append((int(p.style.name[-1]), p.text.strip()))
    def entry(level, text, page, first=False, last=False):
        tabs = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs>'
        pre = ''
        if first:
            pre = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        post = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if last else ''
        rpr = '<w:rPr><w:noProof/></w:rPr>'
        return parse_xml('<w:p %s><w:pPr><w:pStyle w:val="TDC%d"/>%s</w:pPr>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r><w:r>%s<w:tab/></w:r><w:r>%s<w:t>%s</w:t></w:r>%s</w:p>' % (
            W, level, tabs, pre, rpr, text.replace('&', '&amp;'), rpr, rpr, page, post))
    anchor = toc_old[0]._p
    new_entries = []
    for k, (lv, t) in enumerate(heads):
        page = roman.get(t) or (pages or {}).get(('h', t)) or ''
        new_entries.append(entry(lv, t, page, first=(k == 0), last=(k == len(heads) - 1)))
    for e in new_entries: anchor.addprevious(e)
    for p in toc_old: p._p.getparent().remove(p._p)

    doc.save(OUT)
    ps = Document(OUT).paragraphs
    j = next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.startswith('Capítulo I:'))
    entries = [('h', p.text.strip()) for p in ps[j:] if p.style.name in ('Heading 1', 'Heading 2', 'Heading 3')]
    return entries, caps_t, caps_f

if __name__ == '__main__':
    e, ct, cf = build(None)
    pdf = to_pdf(OUT); pages, start, n = compute(pdf, e, ct, cf)
    build(pages)
    pdf = to_pdf(OUT); p2, _, n2 = compute(pdf, e, ct, cf)
    dif = {k: (pages[k], p2[k]) for k in pages if p2.get(k) != pages[k]}
    print('páginas', n2, 'diferencias', len(dif), list(dif.items())[:4])
    if dif: build(p2); pdf = to_pdf(OUT)
