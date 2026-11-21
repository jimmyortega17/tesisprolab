import io, copy
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from PIL import Image, ImageDraw

SRC = 'tpl.pptx'
OUT = 'Lienzos_Cosecha_Viva.pptx'
p = Presentation(SRC)
S = {i + 1: s for i, s in enumerate(p.slides)}

def shp(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid: return sh
        if sh.shape_type == 6:
            for g in sh.shapes:
                if g.shape_id == sid: return g
    raise KeyError(sid)

def set_text(sh, text):
    tf = sh.text_frame
    paras = tf.paragraphs
    first = None
    for para in paras:
        if para.runs: first = para; break
    if first is None:
        first = paras[0]; first.add_run()
    runs = first.runs
    runs[0].text = text
    for r in runs[1:]: r._r.getparent().remove(r._r)
    for para in paras:
        if para is not first and para._p.getparent() is not None: para._p.getparent().remove(para._p)

def set_cell(cell, text):
    tf = cell.text_frame
    para = tf.paragraphs[0]
    if para.runs:
        para.runs[0].text = text
        for r in para.runs[1:]: r._r.getparent().remove(r._r)
    else:
        para.add_run().text = text
    for extra in tf.paragraphs[1:]: extra._p.getparent().remove(extra._p)

COLORS = ['3F7D3A', 'C8762B', '6B4226', '2D6A8F', 'B8332B', '2F4A2B']
def tile(icon, color, size=300):
    base = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(base)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=int(size * 0.18), fill='#' + color)
    ic = Image.open(f'icons/{icon}.png').convert('RGBA').resize((int(size * 0.56), int(size * 0.56)))
    base.paste(ic, ((size - ic.width) // 2, (size - ic.height) // 2), ic)
    b = io.BytesIO(); base.save(b, 'PNG'); return b.getvalue()

def swap_pic(slide, sid, blob, keep_ratio=True, scale=0.82):
    pic = shp(slide, sid)
    _, rid = slide.part.get_or_add_image_part(io.BytesIO(blob))
    blip = pic._element.blipFill.find(qn('a:blip'))
    blip.set(qn('r:embed'), rid)
    # quitar recortes de la imagen original
    src = pic._element.blipFill.find(qn('a:srcRect'))
    if src is not None: pic._element.blipFill.remove(src)
    if keep_ratio:
        side = int(min(pic.width, pic.height) * scale)
        cx, cy = pic.left + pic.width // 2, pic.top + pic.height // 2
        pic.width = pic.height = side; pic.left = cx - side // 2; pic.top = cy - side // 2

def icons(slide, ids, names):
    for i, (sid, n) in enumerate(zip(ids, names)): swap_pic(slide, sid, tile(n, COLORS[i % len(COLORS)]))

# =============== 1. BLANCO DE RELEVANCIA ===============
s = S[1]
for sid, t in {
    337: 'Compra asegurada con prima de 33%',
    338: 'Videos por etapa y QR por lote para ver el origen',
    339: 'La preventa financia la tecnificación del productor',
    350: 'Kit de fertirriego y secado en comodato, sin costo inicial',
    330: 'El precio de S/ 140 aún no se prueba con compra real',
    331: 'Retención y conversión sin evidencia propia',
    332: 'El productor podría vender a otro comprador',
    333: 'Un mal año por clima o plagas reduce el café selecto',
    334: '¿Pagará el consumidor de Lima S/ 140 por adelantado?',
    335: '¿Llegarán las fincas piloto a 22 qq/ha?',
    345: '¿Firmarán 15 de 30 productores el acuerdo?',
    348: '¿Recomprará 60% de los clientes cada año?',
    341: 'Tours a la finca como segundo ingreso',
    343: 'Paquetes de regalo y ventas corporativas',
    347: 'Replicar en Satipo, Villa Rica y otros cultivos',
    349: 'Alianzas con tostadurías y hoteles de la zona',
}.items(): set_text(shp(s, sid), t)

# =============== 2. ¿QUÉ ES? / ¿QUÉ NO ES? ===============
s = S[2]
for sid, t in {
    13: 'Baja productividad: ~15 qq/ha frente a un potencial de 50.',
    14: 'Venta sin diferenciar al intermediario, a precio bajo.',
    15: 'Sin capital ni comprador seguro: no invierte en abono.',
    16: 'Consecuencia social: migración de jóvenes y abandono de fincas.',
    22: 'No es solo un problema de precio: sin tecnificación no hay calidad que pagar.',
    21: 'No es solo un problema de productividad: sin comprador no hay incentivo a invertir.',
    20: 'No es un problema que resuelva el turismo cafetalero por sí solo.',
    19: 'No es un asunto que dependa del Estado: se resuelve con recursos de mercado.',
}.items(): set_text(shp(s, sid), t)
set_text(shp(s, 12), '3. Contexto: Chanchamayo, Junín; parcelas de ~2 ha; consumidor urbano a más de 300 km.')
set_text(shp(s, 11), '4. Solución: tecnifica, compra con prima, vende por adelantado y muestra el origen.')
icons(s, [1026, 1028, 1030, 1032, 1034, 1036, 1038, 1040],
      ['FaSeedling', 'FaStore', 'FaMoneyBillWave', 'FaUsers', 'FaMoneyBillWave', 'FaSeedling', 'FaMapMarkedAlt', 'FaFlag'])

# =============== 3. STORYBOARD DEL PRODUCTOR ===============
s = S[3]
th3 = ['Prepara la parcela, pero abona poco o nada: no tiene dinero para fertilizantes.',
       'Espera que la floración y la lluvia traigan una buena campaña.',
       'Aparecen la roya y la broca y no sabe cómo controlarlas sin técnico.',
       'Cosecha todo junto, maduro y verde: se mezclan las calidades y baja el precio.',
       'Llueve durante el secado en el patio y la humedad daña el grano.',
       'Lleva su café al acopiador más cercano.',
       'El intermediario fija el precio y no puede negociar.',
       'Cobra y siente alivio: cubre las deudas de la campaña.',
       'No le queda dinero para abonar la siguiente campaña.',
       'Teme que sus hijos migren a la ciudad si el ingreso no mejora.']
for sid, t in zip(range(430, 440), th3): set_text(shp(s, sid), t)
icons(s, [2, 3, 4, 11, 1034, 1036, 1038, 1040, 1042, 1044],
      ['FaLeaf', 'FaCloudRain', 'FaExclamationTriangle', 'FaCoffee', 'FaTint', 'FaTruck', 'FaStore', 'FaMoneyBillWave', 'FaHandHoldingUsd', 'FaChild'])

# =============== 4. META-USUARIO ===============
s = S[4]
set_text(shp(s, 2), 'Matriz de META-USUARIO: Rosa')
for sid, t in {
    16: 'Necesita capital para abonar y tecnificar su finca',
    25: 'Vende al acopiador de su zona',
    32: 'Sus vecinos también venden café',
    26: 'edad 42 años',
    27: '',
    21: 'Vive en Chanchamayo, Junín',
    4: 'Casada, con hijos',
    31: 'caficultora',
    18: 'dos hectáreas de café; cosecha ~30 qq por campaña',
    19: 'Desea un mejor precio y un comprador seguro por su café',
    24: 'le agobia que el intermediario fije el precio',
    28: 'espera producir más y mejor café con apoyo técnico',
    29: 'no tiene capital para abonar ni fertirriego; vende sin diferenciar calidades',
    22: 'Anhela que sus hijos estudien y no migren',
    23: 'Orgullosa de su café',
}.items(): set_text(shp(s, sid), t)
swap_pic(s, 8, tile('FaFemale', '3F7D3A', 400), scale=0.95)
for r in shp(s, 4).text_frame.paragraphs[0].runs: r.font.size = Pt(11)

# =============== 5. STORYBOARD DEL CONSUMIDOR ===============
s = S[5]
set_text(shp(s, 53), 'Figura 5   Mapa de experiencia: Carla, consumidora')
th5 = ['Ve en Instagram un video de Rosa explicando su café.',
       'Se emociona: ve el origen del café y a la productora.',
       'Duda: ¿es real el origen? ¿conviene pagar por adelantado?',
       'Lee la garantía de entrega, la fecha de tueste y la ficha del productor, y decide.',
       'Compra por adelantado un paquete de S/ 140 con pago digital.',
       'Espera: recibe por WhatsApp videos de la cosecha, el secado y el tueste.',
       'En junio viaja a Chanchamayo, recorre la finca con Rosa y prueba el café recién cosechado.',
       'Recibe cuatro bolsas con fecha de tueste y un código QR a la ficha de Rosa.',
       'Prepara el café y comparte el video de su productora.',
       'Renueva su paquete y regala otro a su madre.']
for sid, t in zip([430, 431, 432, 433, 434, 435, 436, 437, 438, 439], th5): set_text(shp(s, sid), t)
icons(s, [14, 1026, 8, 6, 9, 10, 13, 1030, 1028, 7],
      ['FaInstagram', 'FaVideo', 'FaBalanceScale', 'FaQrcode', 'FaCreditCard', 'FaWhatsapp', 'FaMapMarkedAlt', 'FaBoxOpen', 'FaCoffee', 'FaGift'])
# emojis: neutral(18) estrella(19) duda(20) feliz(23)
emo = {'neutral': 18, 'star': 19, 'sad': 20, 'happy': 23}
blob_of = {k: shp(S[3], v).image.blob for k, v in emo.items()}
seq = ['neutral', 'star', 'sad', 'happy', 'star', 'neutral', 'star', 'happy', 'happy', 'star']
ids5 = [18, 19, 20, 21, 22, 80, 82, 23, 84, 85]
for sid, e in zip(ids5, seq):
    pic = shp(s, sid)
    _, rid = s.part.get_or_add_image_part(io.BytesIO(blob_of[e]))
    pic._element.blipFill.find(qn('a:blip')).set(qn('r:embed'), rid)
# curva de emociones: puntos y líneas
level = {'neutral': 4.74, 'star': 4.31, 'sad': 4.98, 'happy': 4.45}
dot_ids = [441, 442, 443, 445, 444, 446, 447, 448, 449, 450]
dots = [shp(s, i) for i in dot_ids]
cols = {'neutral': '9E9E9E', 'star': '3F7D3A', 'sad': 'B8332B', 'happy': '6AA84F'}
for d, e in zip(dots, seq):
    d.top = Emu(int(level[e] * 914400))
    d.fill.solid(); d.fill.fore_color.rgb = RGBColor.from_string(cols[e])
for k, lid in enumerate(range(451, 460)):
    ln = shp(s, lid); a, b = dots[k], dots[k + 1]
    ln.begin_x = a.left + a.width; ln.begin_y = a.top + a.height // 2
    ln.end_x = b.left; ln.end_y = b.top + b.height // 2

# =============== 6. LIENZO 6x6 ===============
s = S[6]
tbl = [sh for sh in s.shapes if sh.has_table][0].table
set_cell(tbl.cell(1, 0), 'Los pequeños productores de café de Chanchamayo obtienen ingresos bajos e inestables: producen por debajo de su potencial, venden sin diferenciar a intermediarios y no llegan al consumidor urbano que paga por calidad y origen.')
set_cell(tbl.cell(1, 3), 'El productor necesita producir más y mejor con apoyo técnico, un comprador seguro que pague por calidad y capital antes de la cosecha. El consumidor necesita café fresco, con origen verificable e impacto visible.')
set_cell(tbl.cell(3, 0), 'A partir de la necesidad se plantearon 6 preguntas generadoras, una por necesidad')
qs = ['¿Cómo elevar el rendimiento del productor sin que tenga capital?', '¿Cómo asegurar un comprador que pague por calidad?', '¿Cómo financiar la campaña antes de la cosecha?',
      '¿Cómo dar al consumidor confianza en el origen?', '¿Cómo acercar al consumidor a la finca?', '¿Cómo generar ingresos complementarios al productor?']
ideas = [['Asistencia técnica gratuita', 'Compromiso de compra con prima', 'Preventa al consumidor', 'Videos por etapa', 'Tour en la cosecha', 'Pago por visitante'],
         ['Kit de fertirriego en comodato', 'Venta a cooperativas', 'Adelanto de la empresa', 'Código QR por lote', 'Visitas virtuales', 'Venta de subproductos'],
         ['Crédito en insumos', 'Subasta de microlotes', 'Fondo rotatorio', 'Certificación privada', 'Encuentro con el productor', 'Talleres de catación']]
sel = ['Tecnificar con kit en comodato y asistencia técnica', 'Compra asegurada con prima de 33%', 'Preventa que financia la campaña', 'Videos por etapa y QR por lote', 'Tour guiado a la finca en la cosecha', 'Pago por visitante al productor']
for j in range(6):
    set_cell(tbl.cell(4, j), qs[j])
    for k in range(3): set_cell(tbl.cell(5 + k, j), ideas[k][j])
    set_cell(tbl.cell(11, j), sel[j])
# quitar filas sobrantes (8 a 10)
trs = tbl._tbl.findall(qn('a:tr'))
for i in (8, 9, 10): tbl._tbl.remove(trs[i])

# =============== 7. MATRIZ COSTO / IMPACTO ===============
s = S[7]
for sid, t in {
    14: 'D. Tecnificación + compra asegurada + preventa con trazabilidad + tour (impacto 5, costo 3): ELEGIDA.',
    16: 'Es la única que actúa sobre las dos causas del núcleo crítico y se financia con el dinero del cliente.',
    20: 'B. Asistencia técnica y venta a cooperativas (impacto 3, costo 3): descartada, no cambia el precio.',
    15: 'A. Turismo experiencial cafetalero (impacto 2, costo 2): se conserva como complemento.',
    2: 'C. Tienda en línea de café de Chanchamayo (impacto 2, costo 2): descartada, no tecnifica ni asegura compra.',
    17: 'Ninguna alternativa cae en este cuadrante: alto costo y bajo impacto.',
}.items(): set_text(shp(s, sid), t)
icons(s, [3074, 3076, 3078, 3080, 3084, 3082], ['FaSeedling', 'FaHandshake', 'FaUsers', 'FaMapMarkedAlt', 'FaStore', 'FaExclamationTriangle'])

# =============== 8. LIENZO DE PROPUESTA DE VALOR ===============
s = S[8]
for sid, t in {
    60: 'Productor: asistencia técnica, kit de fertirriego y secador solar en comodato, compra del café selecto a S/ 20/kg. Consumidor: paquete de cosecha de 1 kg a S/ 140, videos por etapa, QR por lote y tour a la finca.',
    65: 'Rendimiento de 15 a 25 qq/ha; ingreso neto del productor +64.7%; pago adicional por visitante; relación con un productor concreto y reporte de impacto.',
    67: 'Kit sin costo inicial; compra asegurada con prima de 33% y piso garantizado; preventa que financia la campaña; videos con su nombre y finca; tueste por lote con fecha en la bolsa.',
    82: 'Producir más y mejor, vender a un comprador seguro con prima y recibir apoyo técnico sin endeudarse; café fresco y trazable con impacto visible.',
    84: 'Vivir de su café y sostener a su familia; vender la cosecha y pagar la campaña; tomar un buen café, regalar algo con historia y apoyar a un productor.',
    88: 'Bajo rendimiento; precio bajo y volátil; falta de capital; trabajo poco valorado; el consumidor no sabe de dónde viene su café ni cuán fresco es.',
}.items(): set_text(shp(s, sid), t)
swap_pic(s, 4, tile('FaBoxOpen', '6B4226', 400), scale=0.7)
swap_pic(s, 3, tile('FaUsers', '2D6A8F', 400), scale=0.7)

# slide 6: más alto y con letra legible
for i, r in enumerate(tbl.rows):
    if i >= 4:
        r.height = Emu(int(0.62 * 914400)) if i != 4 else Emu(int(0.6 * 914400))
        for c in r.cells:
            for para in c.text_frame.paragraphs:
                for run in para.runs: run.font.size = Pt(11)
p.save(OUT)
print('guardado')
