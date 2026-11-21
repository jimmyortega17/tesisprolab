import copy, re, subprocess, os, sys, shutil
from docx import Document
from docx.shared import Inches
from fb_edit import actualizar_cifras, set_par, find_par
from fb_insert import FDoc, move_groups, run as mkrun, para
from contenido1 import parte1
from contenido2 import parte2

OUT = '../Tesis_ProLab_Cosecha_Viva_FINAL.docx'
SCR = '/tmp/claude-0/-home-user-tesisprolab/45d5a49c-f407-5ad1-bffb-270b39202e9f/scratchpad/'

def heading(doc, text, style):
    for p in doc.paragraphs:
        if p.style.name == style and p.text.strip() == text: return p
    raise KeyError(text)

def build(pages=None):
    doc = Document('tesis_original.docx')
    actualizar_cifras(doc)
    # texto de la introducción de 6.3 (el recupero de capital de trabajo ahora es nulo)
    for p in doc.paragraphs:
        if p.text.startswith('La viabilidad se evalúa con un flujo'):
            for r in p.runs: r.text = r.text.replace('con la recuperación del capital de trabajo y el valor de rescate de los activos', 'con el valor de rescate de los activos (el capital de trabajo que se recupera es nulo en el caso base)')
    F = FDoc(doc); parte1(F); parte2(F)
    FIG = '../figuras/'
    def fig(key, texto, titulo, archivo):
        F.at(key); n = F.nf + 1
        F.p(texto.format(n=n)); F.figure(titulo, FIG + archivo, note='Elaboración propia con la plantilla de lienzos de la consultora de empresas.', width=5.6)
    F.at('fca'); n_ = F.nt + 1
    F.p(f'La Tabla {n_} muestra la relación entre el FCL y el FCA. Como el proyecto no usa préstamo, no hay préstamo, amortización, intereses ni escudo fiscal, y el FCA es igual al FCL.')
    from fb_data import FW, n0
    def fila(lab, r): return [lab] + [n0(FW.cell(r, 2 + i).value) for i in range(6)]
    F.table('Flujo de Caja Libre (FCL) y Flujo de Caja del Accionista (FCA) (S/)', ['Concepto', 'Año 0', 'Año 1', 'Año 2', 'Año 3', 'Año 4', 'Año 5'],
            [fila('FCL', 17), fila('(+) Préstamo', 18), fila('(-) Amortización', 20), fila('(-) Intereses', 21), fila('(+) Escudo fiscal', 22), fila('FCA', 23)],
            note='Adaptado del modelo financiero (hoja FCA_WACC). El FCL se descuenta con el WACC y el FCA con el COK; sin préstamo, ambos son iguales y el WACC es igual al COK.', widths=[1.9, 0.9, 0.9, 0.9, 0.9, 0.9, 0.9])
    fig('l1', 'La Figura {n} presenta el lienzo blanco de relevancia: lo que resulta interesante de la solución, las críticas constructivas, las nuevas preguntas y las nuevas ideas.', 'Lienzo blanco de relevancia de Cosecha Viva', 'lienzo_1_blanco_relevancia.png')
    fig('l1', 'La Figura {n} delimita el problema con el lienzo ¿Qué es? y ¿Qué no es?, que separa lo que es el problema de lo que no es.', 'Lienzo ¿Qué es? y ¿Qué no es? del problema', 'lienzo_2_que_es_que_no_es.png')
    fig('l5', 'La Figura {n} resume al usuario-beneficiario en la matriz de meta-usuario de Rosa, una productora de café de Chanchamayo.', 'Matriz de meta-usuario: Rosa, productora de café', 'lienzo_4_meta_usuario.png')
    fig('l8', 'La Figura {n} presenta el storyboard del productor en la situación actual, con sus pensamientos, emociones y la curva antes, durante y después.', 'Storyboard del productor en la situación actual', 'lienzo_3_storyboard_productor.png')
    fig('c3b', 'La Figura {n} muestra el storyboard del consumidor con Cosecha Viva, con una curva emocional que sube a medida que ve el origen de su café.', 'Storyboard del consumidor con Cosecha Viva', 'lienzo_5_storyboard_consumidor.png')
    fig('l9', 'La Figura {n} reproduce el lienzo 6x6 con las preguntas generadoras, sus ideas y la idea elegida para cada una.', 'Lienzo 6x6: preguntas generadoras e ideas', 'lienzo_6_lienzo_6x6.png')
    fig('l10', 'La Figura {n} ubica las alternativas en la matriz costo-impacto; la alternativa D es la elegida.', 'Matriz costo-impacto de las alternativas', 'lienzo_7_costo_impacto.png')
    fig('l12', 'La Figura {n} presenta el lienzo de la propuesta de valor, con el mapa de valor de Cosecha Viva frente al perfil del usuario.', 'Lienzo de la propuesta de valor', 'lienzo_8_propuesta_valor.png')
    anchors = {
        'c2': heading(doc, 'Capítulo III: Investigación del Usuario', 'Heading 1')._p,
        'c3a': heading(doc, '3.2 Mapa de Experiencia de Usuario', 'Heading 2')._p,
        'c3b': heading(doc, '3.3 Identificación de la Necesidad', 'Heading 2')._p,
        'c4': heading(doc, '4.5 Producto Mínimo Viable (PMV)', 'Heading 2')._p,
        'c5a': heading(doc, '5.4 Sostenibilidad del Modelo de Negocio', 'Heading 2')._p,
        'c5b': heading(doc, 'Capítulo VI: Validación del Modelo de Negocio', 'Heading 1')._p,
        'c6a': heading(doc, '6.2 Validación de la Factibilidad de la Solución', 'Heading 2')._p,
        'c6b': heading(doc, 'Capítulo VII: Solución Sostenible', 'Heading 1')._p,
        'c7': heading(doc, 'Capítulo VIII: Decisión e Implementación', 'Heading 1')._p,
        'c8': heading(doc, '8.2 Conclusiones', 'Heading 2')._p,
        'fca': heading(doc, '6.3.7 Indicadores de Rentabilidad y Análisis Costo-Beneficio', 'Heading 3')._p,
        'l1': find_par(doc, 'El lienzo muestra que actuar solo')._p,
        'l5': find_par(doc, 'Los resultados de campo orientan el perfil')._p,
        'l8': find_par(doc, 'El mapa muestra que los momentos críticos')._p,
        'l9': find_par(doc, 'Las ideas se agruparon en cuatro alternativas')._p,
        'l10': find_par(doc, 'La alternativa D se eligió')._p,
        'l12': find_par(doc, 'El elemento central de la propuesta')._p,
        'END': None}
    # los grupos sin clave (introducción) se eliminan; 'END' se queda al final
    keep = {k: v for k, v in anchors.items() if k != 'END'}
    move_groups(F, keep)
    # elementos huérfanos (creados antes del primer at) -> quitar
    body = doc.element.body
    sect = body[-1]
    # ---------- referencias nuevas
    refs = [p for p in doc.paragraphs if p.style.name == 'Normal' and p.paragraph_format.left_indent is not None and p.text.startswith(('Banco Central', 'Osterwalder, A., & Pigneur'))]
    tpl = refs[0]._p
    def newref(parts):
        el = copy.deepcopy(tpl)
        for r in list(el.findall('.//{*}r')): el.remove(r)
        from docx.oxml import parse_xml
        from fb_insert import W, FONT, esc
        for t, it in parts:
            el.append(parse_xml(f'<w:r {W}><w:rPr>{FONT}{"<w:i/><w:iCs/>" if it else ""}<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>'))
        return el
    bn = refs[0]._p
    bn.addnext(newref([('Bland, D. J., & Osterwalder, A. (2019). ', 0), ('Testing business ideas: A field guide for rapid experimentation', 1), ('. John Wiley & Sons.', 0)]))
    refs[1]._p.addnext(newref([('Osterwalder, A., Pigneur, Y., Bernarda, G., & Smith, A. (2014). ', 0), ('Value proposition design: How to create products and services customers want', 1), ('. John Wiley & Sons.', 0)]))
    # ---------- renumeración de tablas y figuras en orden de documento
    cap = re.compile(r'^(Tabla|Figura) (\d+)$')
    seq = {'Tabla': 0, 'Figura': 0}; mapa = {}
    for p in doc.paragraphs:
        m = cap.match(p.text.strip())
        if m and p.runs and p.runs[0].bold:
            seq[m.group(1)] += 1; mapa[(m.group(1), int(m.group(2)))] = seq[m.group(1)]
    def sub(text):
        def one(m): 
            k = (m.group(1), int(m.group(3))); return f'{m.group(1)}{m.group(2)} {mapa.get(k, int(m.group(3)))}'
        def two(m):
            return f'Tablas {mapa.get(("Tabla", int(m.group(1))), int(m.group(1)))} y {mapa.get(("Tabla", int(m.group(2))), int(m.group(2)))}'
        t = re.sub(r'Tablas (\d+) y (\d+)', two, text)
        return re.sub(r'\b(Tabla|Figura)(s?) (\d+)\b', lambda m: (one(m) if m.group(2) == '' else m.group(0)), t)
    lista_ini = {p._p for p in doc.paragraphs[:140] if re.match(r'^(Tabla|Figura) \d+\. ', p.text)}
    def fix(par):
        if par._p in lista_ini: return
        for r in par.runs:
            t = sub(r.text)
            if t != r.text: r.text = t
    for p in doc.paragraphs: fix(p)
    for t in doc.tables:
        for rw in t.rows:
            for c in rw.cells:
                for p in c.paragraphs: fix(p)
    # ---------- índices y listas
    rebuild_indices(doc, mapa, pages)
    doc.save(OUT)
    return doc, mapa

def rebuild_indices(doc, mapa, pages):
    ps = doc.paragraphs
    idx = {p.text.strip(): i for i, p in enumerate(ps[:140]) if p.text.strip() in ('Índice de Contenidos', 'Lista de Tablas', 'Lista de Figuras')}
    i_toc, i_lt, i_lf = idx['Índice de Contenidos'], idx['Lista de Tablas'], idx['Lista de Figuras']
    tpl_ch = copy.deepcopy(ps[23]._p); tpl_sub = copy.deepcopy(ps[24]._p); tpl_ap = copy.deepcopy(ps[58]._p); tpl_li = copy.deepcopy(ps[67]._p)
    # fin de la lista de figuras: último párrafo 'Figura n.' antes del primer bloque de contenido
    j = i_lf + 1
    while re.match(r'^Figura \d+\. ', ps[j].text): j += 1
    first_body = ps[j]._p
    old_toc = [p._p for p in ps[i_toc + 1:i_lt]]; old_lt = [p._p for p in ps[i_lt + 1:i_lf]]; old_lf = [p._p for p in ps[i_lf + 1:j]]
    # recopilar entradas desde el cuerpo
    body_ps = ps[j:]
    entries = []   # (nivel, texto)
    ap_started = False
    for p in body_ps:
        s = p.style.name; t = p.text.strip()
        if s == 'Heading 1': entries.append(('ch', t)); ap_started = (t == 'Apéndices')
        elif s == 'Heading 2': entries.append(('ap' if ap_started else 'sub', t))
    caps_t = []; caps_f = []
    lst = body_ps
    for k, p in enumerate(lst):
        m = re.match(r'^(Tabla|Figura) (\d+)$', p.text.strip())
        if m and p.runs and p.runs[0].bold:
            (caps_t if m.group(1) == 'Tabla' else caps_f).append((int(m.group(2)), lst[k + 1].text.strip()))
    def pg(kind, key):
        return (pages or {}).get((kind, key), '')
    def mk(tpl, text, page, bold_first=True):
        el = copy.deepcopy(tpl); rs = el.findall('.//{*}r')
        rs[0].find('{*}t').text = text; rs[1].find('{*}t').text = f'\t{page}'
        return el
    anchor_toc = old_toc[0]
    for e in old_toc: e.getparent().remove(e)
    for e in old_lt: e.getparent().remove(e)
    for e in old_lf: e.getparent().remove(e)
    lt_head = ps[i_lt]._p; lf_head = ps[i_lf]._p
    for kind, t in entries:
        el = mk({'ch': tpl_ch, 'sub': tpl_sub, 'ap': tpl_ap}[kind], t, pg('h', t))
        lt_head.addprevious(el)
    for n, t in caps_t: lf_head.addprevious(mk(tpl_li, f'Tabla {n}. {t}', pg('Tabla', n)))
    for n, t in caps_f: first_body.addprevious(mk(tpl_li, f'Figura {n}. {t}', pg('Figura', n)))
    return entries, caps_t, caps_f

if __name__ == '__main__':
    build()
    print('ok')
