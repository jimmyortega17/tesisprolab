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

    # utilidades extra
    def h2s(text): return clone_heading(H2, text)
    def runin(like, text):
        el = clone_par(like, text)
        for r in el.iter(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is None: rpr = parse_xml('<w:rPr %s/>' % W); r.insert(0, rpr)
            if rpr.find(qn('w:b')) is None: rpr.insert(0, parse_xml('<w:b %s/>' % W))
        return el
    def label(text): return runin(TXT, text)
    F = FDoc(doc); F.nt = 99; F.nf = 98
    hdr = ['Concepto', 'Año 1', 'Año 2', 'Año 3', 'Año 4', 'Año 5']

    # ================= Capítulo I =================
    B[213].addprevious(h3('1.2.1 Lienzo de Dos Dimensiones: Delimitación del Problema'))
    p1 = clone_par(218, 'Para precisar el problema se aplica el Lienzo de Dos Dimensiones, que distingue lo que forma parte del problema (qué es) y lo que queda fuera de su alcance (qué no es). La Figura 1 presenta el lienzo.')
    p2 = txt('Qué es: la baja productividad (cerca de 15 qq/ha frente a un potencial de 50), la venta de café sin diferenciar al intermediario a un precio bajo y volátil, la falta de capital y de un comprador seguro, y la migración de jóvenes y el abandono de fincas.')
    p3 = txt('Qué no es: no es solo un problema de precio, porque sin tecnificación no hay calidad que pagar; no es solo un problema de productividad, porque sin comprador no hay incentivo a invertir; no es un problema que resuelva el turismo cafetalero por sí solo; y no depende del Estado, porque se resuelve con recursos de mercado.')
    B[218].addprevious(p1); B[218].addprevious(p2); B[218].addprevious(p3); rm([B[218]])
    rm([B[225], B[233]])                      # 1.3.1 y 1.3.2 no existen en el modelo

    # ================= Capítulo II =================
    seg = B[267:273]; pestel = B[255:260]; legal = B[475:480]
    B[239].addprevious(h3('2.1.1 Dimensionamiento del Mercado: Embudo TAM-SAM-SOM'))
    B[245].addprevious(h3('2.1.2 Respaldo Cualitativo Exploratorio del Dimensionamiento'))
    before(B[246], seg)
    rm([B[266], B[254], B[260], B[265]])
    por_intro = clone_par(261, 'Las cinco fuerzas de Porter se evaluaron para el negocio de café de especialidad con preventa en Lima. La Tabla 6 resume la intensidad de cada fuerza y la respuesta del modelo. La valoración corresponde al juicio del equipo y debe contrastarse con entrevistas a tostadurías y cafeterías.')
    rm([B[261]])
    intro22 = txt('El análisis competitivo se realiza con las cinco fuerzas de Porter y con la comparación de competidores directos e indirectos.')
    opor = [txt('Las principales oportunidades son: (a) la ausencia, según la búsqueda competitiva documentada (Apéndice G), de una oferta en el Perú que combine tecnificación del productor, compra asegurada, venta anticipada y trazabilidad audiovisual; (b) el crecimiento del consumo interno de café, que alcanzaría 1.7 kilogramos por persona al año (Agraria.pe, 2025), y de las cafeterías y marcas de especialidad de Lima (Perú21, s.f.); y (c) los precios históricos del café (Agraria.pe, 2026), que mejoran la capacidad de pago de la cadena.'),
            txt('Los riesgos identificados son: (a) el alza del precio del café en chacra, que puede dejar el precio pactado por debajo del mercado; (b) la venta del productor a otro comprador; (c) el clima y las plagas, como la roya y la broca, que reducen el volumen y la calidad; y (d) que el consumidor de Lima no acepte pagar por adelantado el precio de S/ 140, supuesto que se valida en el piloto. Las medidas de mitigación se presentan en el capítulo VI y el análisis completo del entorno en el Apéndice I.')]
    vent = [txt('La tienda en línea y el tueste con aliados son fáciles de replicar; lo difícil de replicar es la combinación de cuatro elementos. Primero, la red de productores tecnificados con acuerdos escritos, que exige presencia de campo y tiempo, porque la tecnificación necesita una campaña para madurar. Segundo, el kit en comodato y la asistencia técnica, que crean un vínculo con el productor y que un nuevo entrante tendría que financiar antes de vender.'),
            txt('Tercero, la base de clientes y los datos de cada lote, que sostienen la relación durante todo el año y la recompra. Cuarto, la marca ligada a fincas y productores con nombre, que se construye con videos y visitas. La defensa depende de sostener la prima con piso garantizado, de mantener una red de productores de respaldo y de medir la retención en el piloto; las barreras de entrada son medias y no impiden que un competidor con más capital intente replicar el modelo.')]
    seq = [intro22, h3('2.2.1 Fuerzas Competitivas: Análisis de Cinco Fuerzas de Porter'), por_intro, B[262], B[263], B[264], B[252],
           h3('2.2.2 Mapa de Posicionamiento Estratégico y Análisis Competitivo'), B[247], B[248], B[249], B[250], B[251], B[253],
           h3('2.2.3 Oportunidades y Riesgos del Entorno Competitivo')] + opor + [h3('2.2.4 Ventajas Competitivas Sostenibles y Barreras de Entrada Dinámicas')] + vent
    before(B[273], seq)

    # ================= Capítulo III =================
    meta_fig = B[282:287]; empat = B[300:311]
    rm([B[299]])
    B[311].addprevious(txt('El lienzo de meta-usuario de Rosa y los mapas de empatía del productor y del consumidor se presentan en el Apéndice J.'))
    rm([B[323]])
    F.at('t34')
    F.p('El perfil y el mapa de experiencia parten de supuestos sobre cómo responderán el consumidor y el productor a la solución. Reconocer cada supuesto, junto con la conducta alternativa que podría invalidarlo, evita asumir que la propuesta se aceptará por sí sola. La Tabla 102 resume los supuestos conductuales y su evidencia requerida.')
    sup = [['El consumidor de Lima paga por adelantado un paquete de S/ 140', 'Prefiere comprar café ya tostado en una tienda, sin esperar la cosecha', 'Conversión menor que 3% y menos de 300 paquetes (H5)', 'Preventa con pago real en el piloto'],
           ['El productor firma el compromiso y entrega su café selecto', 'Vende a otro comprador cuando sube el precio en chacra', 'Falta de café para cumplir la preventa (H6)', '15 de 30 acuerdos firmados y entregas del piloto'],
           ['El consumidor valora la trazabilidad con videos y QR', 'Compra por precio o por marca y no mira los videos', 'La trazabilidad no justifica el precio (H7)', 'Encuesta a 150 compradores y prueba con y sin videos'],
           ['El cliente recompra y visita la finca', 'Compra una vez como regalo y no vuelve', 'Retención menor que 60% y menos de 5% de reservas de tour (H8)', 'Seguimiento de los compradores durante seis meses'],
           ['La tecnificación eleva el rendimiento a 22 qq/ha o más', 'El productor aplica solo parte del plan o hay una mala campaña', 'Más productores por paquete vendido (H4)', 'Medición en diez fincas piloto']]
    F.table('Supuestos Conductuales y Comportamientos Alternativos que Pueden Invalidar la Propuesta', ['Supuesto actual', 'Conducta alternativa posible', 'Riesgo para Cosecha Viva', 'Evidencia requerida'], sup, note='Adaptado de las hipótesis H4 a H8 y de los riesgos del capítulo VI.', widths=[1.7, 1.7, 1.6, 1.5])
    F.p('Estos supuestos orientaron el diseño de las hipótesis y los experimentos del capítulo VI: la validación de campo respaldó el interés (H1 a H3), y los supuestos de compra, compromiso, trazabilidad, recompra y rendimiento se contrastan en el piloto.')
    t34 = [h2s('3.4 Comportamientos que Pueden Invalidar la Propuesta')] + F.groups['t34']
    before(B[338], t34)

    # ================= Capítulo IV =================
    fig5 = B[345:350]; fig6 = B[355:360]; fig7 = B[379:384]; vpc = B[386:397]
    rm([B[385], B[398], B[400], B[407]])
    B[350].addprevious(txt('El lienzo 6x6 se presenta en el Apéndice J.'))
    B[360].addprevious(txt('La matriz costo-impacto se presenta en el Apéndice J.'))
    B[397].addprevious(txt('El lienzo de la propuesta de valor del productor y del consumidor se presenta en el Apéndice J.'))

    # ================= Capítulo V =================
    rm([B[428], B[437], B[474]])
    mk_new = []
    for i in range(438, 474):
        e = B[i]
        if e.tag.endswith('}p') and P(i).style.name == 'Heading 3':
            lab = re.sub(r'^5\.5\.\d+\s+', '', P(i).text.strip()) + '.'
            mk_new.append(label(lab)); rm([e])
        else: mk_new.append(e)
    sigi = [h2s('5.4 Sistema Integral de Gestión de la Innovación'),
            txt('Cosecha Viva gestiona la innovación como un proceso permanente y basado en evidencia. Cada hipótesis tiene un experimento, una métrica y un umbral definidos antes de su aplicación (capítulo VI), y cada resultado se registra en una tarjeta de aprendizaje (Apéndice K) que decide si se escala, se ajusta o se detiene el modelo.'),
            txt('El sistema se organiza en tres niveles. En el nivel estratégico se controlan el VAN financiero, el VAN social, el avance de la escala (clientes activos, paquetes y productores de cada año) y la decisión de escalar después del piloto. En el nivel táctico se controlan las hipótesis críticas H4 a H8 y los indicadores de marketing (visitas calificadas, conversión de 3%, CAC de S/ 80, retención de 60% y relación LTV/CAC mayor que 3; Tabla 28). En el nivel operativo se controlan el rendimiento por hectárea de las fincas piloto, la calidad del café selecto, las entregas a tiempo y la satisfacción del cliente.'),
            txt('Los responsables son los definidos para el piloto (gerente general, responsable comercial, responsable de operaciones, jefe agronómico y socios) y los umbrales de decisión son los de los hitos del piloto (Tabla 66). Si un criterio no se cumple se aplica la acción prevista para ese hito, y los supuestos del modelo financiero, las probabilidades de los escenarios y la simulación Monte Carlo se actualizan con los resultados medidos.')]
    before(B[435], sigi)

    # ================= Capítulo VI =================
    prior = B[497:506]; diseno = B[507:512]; tareas = B[548:553]; tarjetas = B[513:547]; embudo = B[490:496]
    rm([B[496], B[506], B[512], B[547]])
    F.at('sint')
    F.p('La validación de campo (Fase 1) contrastó las hipótesis H1 a H3 con 20 productores y 25 turistas potenciales; las hipótesis H4 a H8 son las que determinan la aceptación del producto definitivo y se contrastan en el piloto. La Tabla 103 resume el cuadro de mando de las ocho hipótesis frente a su criterio de éxito.')
    sin = [['H1. El consumidor valora que su pago beneficie al productor', '70%', '92% (23 de 25)', 'Sí'], ['H2. Paga S/ 100 o más por una experiencia', '60%', '76% (19 de 25)', 'Sí'], ['H3. El productor acepta recibir visitantes', '70%', '85% (17 de 20)', 'Sí'],
           ['H4. La tecnificación eleva el rendimiento a 22 qq/ha o más', '22 qq/ha', 'Pendiente: piloto', 'Por medir'], ['H5. El consumidor de Lima compra por adelantado un paquete', '3% y 300 paquetes', 'Pendiente: piloto', 'Por medir'],
           ['H6. El productor firma el compromiso de venta', '50% (15 de 30)', 'Pendiente: piloto', 'Por medir'], ['H7. El consumidor valora la trazabilidad con videos', '60%', 'Pendiente: piloto', 'Por medir'], ['H8. El cliente recompra y visita la finca', '60% y 5%', 'Pendiente: piloto', 'Por medir']]
    F.table('Resumen de los Resultados de la Fase 1 por Hipótesis', ['Hipótesis', 'Criterio de éxito', 'Resultado medido', '¿Cumple?'], sin, note='Adaptado del trabajo de campo del equipo (Apéndice A) y del plan de validación del piloto.', widths=[3.0, 1.2, 1.4, 0.9])
    sint = [h3('Síntesis de la Validación de la Fase 1 (Cuadro de Mando de las Hipótesis)')] + F.groups['sint']
    before(B[482], sint)
    before(B[484], [h3('6.1.1 Hipótesis para Validar la Deseabilidad de la Solución')])
    call_k = txt('Las tarjetas de prueba y de aprendizaje de las hipótesis H1 a H8 se presentan en el Apéndice K.')
    res13 = [h3('6.1.3 Resultados del Experimento 1: Validación de Campo del PMV de la Primera Versión'),
             txt('El Experimento 1 fue la encuesta de campo aplicada al PMV de la primera versión (20 productores y 25 turistas potenciales, con muestreo por conveniencia). Sus resultados validaron H1, H2 y H3 y dieron origen a la segunda versión del PMV (sección 4.5).')]
    before(B[488], res13)
    # 6.1.2 antes del diseño de experimentos
    exp_block = [h3('6.1.2 Experimentos Empleados para Validar la Deseabilidad de la Solución')] + diseno + [call_k] + tareas
    before(res13[0], prior + exp_block)
    rm([B[489]])
    # 6.2
    intro62 = txt('La factibilidad se evalúa con el plan de mercadeo, el plan de operaciones, las simulaciones de abastecimiento y el piloto de preventa.')
    emb_label = label('Embudo comercial y economía del cliente.')
    sim62 = [h3('6.2.3 Simulaciones para Validar las Hipótesis de Factibilidad'),
             txt('La simulación por cohortes calcula cuántos productores se necesitan cada año según el rendimiento por hectárea y la proporción de café selecto, y la matriz de riesgos identifica qué puede interrumpir la preventa y cómo se mitiga. La Tabla 43 presenta el abastecimiento por cohortes y la Tabla 44 los riesgos y sus medidas.')]
    exp2 = [h3('Experimento 2: Piloto de Preventa con Diez Fincas (H4 a H8)'),
            txt('El piloto de seis meses combina diez fincas con kit y asistencia técnica y una preventa de prueba con pago real de 300 paquetes. Valida las hipótesis H4 a H8 con criterios de decisión definidos antes de conocer los resultados: 3% de conversión y 300 paquetes (H5), 15 de 30 acuerdos firmados (H6), 60% de recompra y 5% de reservas de tour (H8) y 22 qq/ha en la segunda campaña (H4). El presupuesto y los hitos del piloto se presentan en el capítulo VIII.')]
    before(B[565], [intro62, h3('6.2.1 Plan de Mercadeo')] + mk_new + [emb_label] + embudo + [h3('6.2.2 Plan de Operaciones'), B[554], B[560], h3('6.2.3 Simulaciones para Validar las Hipótesis de Factibilidad')] + sim62[1:] + B[555:560] + B[561:565] + exp2)
    # 6.3
    rm([B[567], B[574], B[581], B[588], B[595], B[601], B[612], B[619], B[625], B[637], B[676]])
    def lab(t): return label(t)
    seq63 = [h3('6.3.1 Presupuesto de Inversión')] + B[575:581] + B[683:688] + [lab('Capital de trabajo.')] + B[589:595] + \
            [h3('6.3.2 Análisis Financiero'), lab('Costo de oportunidad del capital (COK).')] + B[568:574] + [lab('Economía unitaria.')] + B[582:588] + \
            [lab('Estado de resultados proyectado.')] + B[596:601] + [lab('Flujo de caja libre.')] + B[602:612] + [lab('Indicadores de rentabilidad y análisis costo-beneficio.')] + B[613:619] + B[677:683] + \
            [lab('Punto de equilibrio.')] + B[620:625] + [h3('6.3.3 Simulaciones Empleadas para Validar las Hipótesis')] + B[626:637] + B[638:676]
    before(B[688], seq63)

    # ================= Capítulo VII =================
    amb = B[710:715]; gen = B[716:721]
    rm([B[709], B[715]])
    before(B[696], amb + gen)
    F.at('ben')
    F.p('Los beneficios sociales son el ingreso bruto adicional de los productores frente a la situación sin el modelo, valorado a precios de mercado. Cada ODS declarado genera un beneficio: el ODS 2 (meta 2.3) genera el valor de la producción adicional que logra el productor al pasar de 15 a 25 quintales por hectárea (20 quintales por productor con tecnificación madura, al precio del intermediario de S/ 690 por quintal); el ODS 1 (meta 1.2) genera el sobreprecio que recibe por su café selecto (S/ 230 por quintal frente al intermediario); y el ODS 8 (metas 8.3 y 8.9) genera la remuneración del empleo calificado de los técnicos de campo y el gasto local de los tours (transporte, guía, almuerzo y pago al productor). La Tabla 100 presenta la proyección.')
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
    lim = [h2s('8.3 Limitaciones del Estudio y Análisis de Escenarios Adversos'),
           txt('El equipo reconoce de forma explícita las limitaciones del estudio, porque declararlas define el alcance de las conclusiones. Se agrupan en cuatro.'),
           txt('Primera limitación (evidencia): la validación de campo se hizo con 20 productores y 25 turistas potenciales, con muestreo por conveniencia, por lo que los porcentajes no se generalizan a todos los caficultores ni a todos los consumidores de Lima. Los 25 encuestados de la demanda son turistas potenciales y no necesariamente compradores recurrentes, y la evidencia mide interés declarado, no compra real. Tampoco existe una línea base propia de rendimiento de los productores que participarán.'),
           txt('Segunda limitación (supuestos de baja confianza): la retención de 60%, el costo de adquisición de S/ 80, la conversión de 3%, el costo adicional de S/ 2,500 por hectárea y los anticipos de clientes no tienen evidencia propia y se miden en el piloto. El precio de S/ 140 es una decisión del equipo que aún no se prueba con compra real.'),
           txt('Tercera limitación (modelo): las variables de la simulación Monte Carlo se muestrearon de forma independiente, aunque en la práctica algunas se mueven juntas; el COK no se recalcula si se incorpora deuda; y el VAN social usa supuestos regionales del productor y no aplica precios sociales a la mano de obra ni a las divisas.'),
           txt('Cuarta limitación (escenarios adversos): el VAN se anula con un precio de S/ 104, es decir, 25.6% por debajo de S/ 140. Los escenarios pesimista y de estrés mantienen un VAN positivo (S/ 616,377 y S/ 98,728), pero el de estrés queda cerca de cero, y la simulación Monte Carlo estima una probabilidad de VAN negativo de 0.4%. Estos resultados dependen de que el consumidor acepte el precio y el volumen de la preventa, por lo que no debe invertirse a escala antes de validarlos en el piloto.')]
    before(B[754], lim)

    # ================= Apéndices nuevos =================
    sect = body[-1]
    def add(els):
        for e in els: sect.addprevious(e)
    add([h2('Apéndice I. Análisis PESTEL Completo del Mercado Peruano'), clone_par(851, 'Este apéndice presenta el análisis del entorno político, económico, social, tecnológico, ecológico y legal en el que opera Cosecha Viva, y el marco legal y regulatorio del modelo.')] + pestel + legal)
    add([h2('Apéndice J. Lienzos: Meta Usuario, Propuesta de Valor, Lienzo 6x6 y Costo vs Impacto'), clone_par(851, 'Este apéndice reúne los lienzos elaborados durante el proceso de tesis.'),
         h3('J.1 Lienzo de Meta Usuario')] + meta_fig + empat + [h3('J.2 Lienzo de Propuesta de Valor')] + fig7 + vpc + [h3('J.3 Lienzo 6x6')] + fig5 + [h3('J.4 Lienzo Costo vs Impacto')] + fig6)
    add([h2('Apéndice K. Tarjetas de Prueba y Aprendizaje'), clone_par(851, 'Este apéndice presenta las tarjetas de prueba y de aprendizaje de las hipótesis H1 a H8.')] + tarjetas)

    # ================= títulos exactos del modelo (AnsioZen) =================
    ren = {'Capítulo I: Definición del Problema': 'Capítulo I. Definición del Problema', 'Capítulo II: Análisis del Mercado': 'Capítulo II. Análisis del Mercado',
           'Capítulo III: Investigación del Usuario': 'Capítulo III. Investigación del Usuario', 'Capítulo IV: Diseño del Producto o Servicio': 'Capítulo IV. Diseño del Producto o Servicio',
           'Capítulo V: Modelo de Negocio': 'Capítulo V. Modelo de Negocio', 'Capítulo VI: Validación del Modelo de Negocio': 'Capítulo VI. Solución Deseable, Factible y Viable',
           'Capítulo VII: Solución Sostenible': 'Capítulo VII. Solución Sostenible', 'Capítulo VIII: Decisión e Implementación': 'Capítulo VIII. Decisión e Implementación',
           '1.1 Contexto en el que se Determina el Problema': '1.1 Contexto del Problema por Resolver',
           '1.3 Sustento de la Complejidad y Relevancia del Problema': '1.3 Sustento de la Complejidad y Relevancia del Problema a Resolver',
           '5.3 Escalabilidad y Exponencialidad del Modelo de Negocio': '5.3 Escalabilidad/Exponencialidad del Modelo de Negocio',
           '5.4 Sostenibilidad del Modelo de Negocio': '5.5 Sostenibilidad del Modelo de Negocio', '8.3 Recomendaciones': '8.4 Recomendaciones'}
    for p in doc.paragraphs:
        if p.style.name.startswith('Heading') and p.text.strip() in ren:
            p.runs[0].text = ren[p.text.strip()]
            for r_ in p.runs[1:]: r_._r.getparent().remove(r_._r)

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
    j = next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.startswith('Capítulo I.'))
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
