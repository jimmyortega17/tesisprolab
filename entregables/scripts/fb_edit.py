import re
from fb_data import *

def set_par(p, text):
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]: r._r.getparent().remove(r._r)

def find_par(doc, start, nth=0):
    hits = [p for p in doc.paragraphs if p.text.startswith(start)]
    assert hits, start
    return hits[nth]

def setc(tbl, r, c, text):
    cell = tbl.rows[r].cells[c]; p = cell.paragraphs[0]
    p.runs[0].text = text
    for x in p.runs[1:]: x._r.getparent().remove(x._r)

def fill(tbl, label_map, ncols, fmt=n0, c_off=1):
    """label_map: {texto inicial de la fila: lista de valores}"""
    for i, r in enumerate(tbl.rows):
        lab = r.cells[0].text.strip()
        for k, vals in label_map.items():
            if lab.startswith(k):
                for j, v in enumerate(vals): setc(tbl, i, c_off + j, fmt(v) if not isinstance(v, str) else v)

def actualizar_cifras(doc):
    T = doc.tables
    base = FCL['B23'].value
    # ---- Tabla 16 (índice 15): LTV y LTV/CAC
    ltv = DEM['C14'].value; lc = DEM['C15'].value
    fill(T[15], {'Valor del cliente': [f'{ltv:,.0f}'] * 5, 'Relación LTV': [f'{lc:.2f}'] * 5}, 5, fmt=lambda x: x)
    # ---- Tabla 21 (idx 20): economía unitaria
    names = ['Valor de venta', 'Compra de café', 'Pilado', 'Maquila', 'Empaque', 'Envío', 'Comisión', 'Costo variable', 'Margen de contribución']
    for k, r in zip(names, [5, 7, 8, 9, 10, 11, 12, 13, 14]):
        for i, rw in enumerate(T[20].rows):
            if rw.cells[0].text.startswith(k):
                setc(T[20], i, 1, n2(UNI.cell(r, 2).value)); setc(T[20], i, 2, pc(UNI.cell(r, 3).value))
    # ---- Tabla 22 (idx 21): capital de trabajo
    fill(T[21], {'Inventario de café': row(CT, 5, 2, 6), 'Inventario de empaques': row(CT, 6, 2, 6), 'Caja mínima': row(CT, 7, 2, 6),
                 '(-) Anticipos': row(CT, 8, 2, 6), 'Capital de trabajo requerido': row(CT, 9, 2, 6)}, 5)
    fill(T[21], {'Capital de trabajo / ventas': [pc(v) for v in row(CT, 10, 2, 6)]}, 5, fmt=lambda x: x)
    # ---- Tabla 23 (idx 22): estado de resultados
    fill(T[22], {'Ventas de paquetes': row(EER, 5, 2, 6), 'Ventas de tours': row(EER, 6, 2, 6), 'Ventas netas': row(EER, 7, 2, 6),
                 '(-) Costo variable del café': row(EER, 8, 2, 6), '(-) Costo variable de tours': row(EER, 9, 2, 6), 'Margen de contribución': row(EER, 10, 2, 6),
                 '(-) Marketing': row(EER, 11, 2, 6), '(-) Costos fijos': row(EER, 12, 2, 6), 'EBITDA': row(EER, 13, 2, 6),
                 '(-) Depreciación': row(EER, 14, 2, 6), 'Utilidad operativa': row(EER, 15, 2, 6), '(-) Impuesto': row(EER, 19, 2, 6), 'Utilidad neta': row(EER, 20, 2, 6)}, 5)
    # ---- Tabla 24 (idx 23): flujo de caja libre
    fill(T[23], {'Ventas netas': row(FCL, 5, 2, 7), 'Recupero de capital': row(FCL, 6, 2, 7), 'Valor de rescate': row(FCL, 7, 2, 7), 'Total ingresos': row(FCL, 8, 2, 7),
                 'Inversión inicial': row(FCL, 9, 2, 7), 'Kits': row(FCL, 10, 2, 7), 'Costos variables': row(FCL, 11, 2, 7), 'Marketing': row(FCL, 12, 2, 7),
                 'Costos fijos': row(FCL, 13, 2, 7), 'Impuesto': row(FCL, 14, 2, 7), 'Inversión en capital': row(FCL, 15, 2, 7), 'Total egresos': row(FCL, 16, 2, 7),
                 'Flujo de caja libre': row(FCL, 17, 2, 7), 'Flujo descontado acumulado': row(FCL, 20, 2, 7)}, 6)
    # ---- Tabla 25 (idx 24): indicadores
    vals = {'VAN financiero al COK': 'S/ ' + n0(base), 'TIR': pc(FCL['B24'].value), 'Periodo de recupero': f"{FCL['B26'].value:.2f} años",
            'Valor presente de los ingresos': 'S/ ' + n0(FCL['B27'].value), 'Valor presente de los egresos': 'S/ ' + n0(FCL['B28'].value),
            'Relación beneficio-costo': f"{FCL['B29'].value:.2f}", 'VAN esperado': 'S/ ' + n0(SEN['B17'].value), 'VAN financiero a la tasa social': 'S/ ' + n0(FCL['B30'].value)}
    for i, rw in enumerate(T[24].rows):
        for k, v in vals.items():
            if rw.cells[0].text.startswith(k): setc(T[24], i, 1, v)
    # ---- Tabla 26 (idx 25): punto de equilibrio
    fill(T[25], {'Paquetes de equilibrio operativo': row(EER, 23, 2, 6), 'Paquetes de equilibrio contable': row(EER, 24, 2, 6)}, 5)
    fill(T[25], {'Margen de seguridad': [pc(v) for v in row(EER, 25, 2, 6)]}, 5, fmt=lambda x: x)
    # ---- Tabla 27 (idx 26): sensibilidad
    for i, r in enumerate(range(5, 10)):
        for j in range(5): setc(T[26], i + 1, 1 + j, n0(SEN.cell(r, 2 + j).value))
    # ---- Tabla 28 (idx 27): escenarios
    for i, r in enumerate(range(13, 18)): setc(T[27], i + 1, 3, n0(SEN.cell(r, 2).value))
    # ---- Tabla 31 (idx 30): VAN social
    for i, (k, c) in enumerate([('(a)', 'B25'), ('(b)', 'B26'), ('(c)', 'B27'), ('VAN social', 'B24')]):
        for j, rw in enumerate(T[30].rows):
            if rw.cells[0].text.startswith(k): setc(T[30], j, 1, n0(PRO[c].value))
    # ---- Tabla 39 (idx 38): supuesto de precio
    for j, rw in enumerate(T[38].rows):
        if rw.cells[0].text.startswith('Precio del paquete de cosecha'):
            setc(T[38], j, 1, '140'); setc(T[38], j, 3, 'Decisión del equipo: S/ 140 para que el flujo del año 1 no sea negativo; referencia de café de especialidad de Chanchamayo a S/ 130 en Mercado Libre Perú (2026)')
    # ---- Tabla 33 (idx 32): hito de preventa
    for j, rw in enumerate(T[32].rows):
        for c, cell in enumerate(rw.cells):
            if 'Probar precio de S/ 110' in cell.text:
                setc(T[32], j, c, 'Probar precio de S/ 130 y luego S/ 120; no escalar si el precio de quiebre (S/ 104) no se sostiene')
    # ---- reemplazos globales simples en párrafos y celdas
    def rep(par):
        for r in par.runs:
            t = r.text
            t = t.replace('S/ 104.40', 'S/ 104').replace('S/ 120', 'S/ 140').replace('alrededor de 30% del valor de venta', 'alrededor de 25% del valor de venta')
            if t != r.text: r.text = t
    for p in doc.paragraphs: rep(p)
    for t in doc.tables:
        for rw in t.rows:
            for c in rw.cells:
                for p in c.paragraphs: rep(p)

    # ---- prosa
    sens = {k: SEN.cell(r, 2).value for k, r in (('12', 28), ('28', 33))}
    vans = PRO['B24'].value; vpriv = PRO['B25'].value; vtax = PRO['B26'].value; vinc = PRO['B27'].value
    m5 = EER['F22'].value
    pe1, pe2 = EER['B23'].value, EER['C23'].value
    ms2 = EER['C25'].value; ms1 = EER['B25'].value
    V = n0
    S = lambda s, new: set_par(find_par(doc, s), new)
    S('Esta tesis propone', "Esta tesis propone Cosecha Viva, una empresa privada que tecnifica al productor con asistencia técnica y un kit de fertirriego y secado en comodato, le compra su café selecto a S/ 20 por kilogramo de pergamino, 33% por encima del intermediario, y lo vende por adelantado a consumidores de Lima y la costa en paquetes de cosecha de un kilogramo a S/ 140. El cliente sigue su café con videos de cada etapa y un código QR por lote, y puede visitar la finca con un tour guiado, que es el segundo ingreso del modelo.")
    S('El trabajo de campo con 20', f"El trabajo de campo con 20 productores y 25 turistas potenciales validó el interés por una experiencia vinculada al café y que 92% valora que su pago beneficie al productor, pero aún no demuestra la compra anticipada del producto definitivo, que se validará en un piloto de diez fincas con criterios de decisión definidos. Con los supuestos actuales y un COK de 20.94%, el proyecto genera un flujo de caja positivo desde el primer año, un VAN de S/ {V(base)}, una TIR de {pc(FCL['B24'].value)} y un VAN esperado por escenarios de S/ {V(SEN['B17'].value)}, sin depender del Estado. El precio sigue siendo la variable crítica: el VAN se anula con S/ 104, pero el precio de S/ 140 todavía no se ha probado con compra real. Una simulación Monte Carlo de 10,000 iteraciones da una probabilidad de VAN negativo de {MCV['p']}. El ingreso neto estimado del productor sube en 64.7% y el VAN social al 8% alcanza S/ {V(vans)}.")
    S('Fieldwork with 20 farmers', f"Fieldwork with 20 farmers and 25 potential tourists validated interest in a coffee-related experience and that 92% value payments that benefit the farmer, but it does not yet prove advance purchases of the final product, which will be tested in a ten-farm pilot with predefined decision criteria. Under current assumptions and a 20.94% cost of equity, the project yields positive free cash flow from the first year, an NPV of S/ {V(base)}, an IRR of {pc(FCL['B24'].value)} and a scenario-weighted expected NPV of S/ {V(SEN['B17'].value)}, with no reliance on the State. Price remains the critical variable: NPV reaches zero at S/ 104, but the S/ 140 price has not yet been tested with real purchases. A 10,000-iteration Monte Carlo simulation gives a {MCV['p']} probability of a negative NPV. Estimated farmer net income rises by 64.7% and the social NPV at 8% reaches S/ {V(vans)}.")
    S('La solución se denomina', "La solución se denomina Cosecha Viva Chanchamayo. Es una empresa privada que tecnifica al pequeño productor, le compra su café selecto a S/ 20 por kilogramo de pergamino, 33% por encima del precio del intermediario, y lo vende por adelantado a consumidores de Lima y la costa en paquetes de cosecha de un kilogramo (cuatro bolsas de 250 gramos) a S/ 140. Mientras su café se produce, el cliente recibe videos de cada etapa y puede visitar la finca con un tour guiado.")
    S('El modelo es viable financieramente', f"El modelo es viable financieramente. Con una inversión inicial de S/ 156,200 y un costo de oportunidad del capital de 20.94%, calculado con el modelo CAPM, beta total y riesgo país, genera, con los supuestos actuales, un VAN de S/ {V(base)} y una TIR de {pc(FCL['B24'].value)} en un horizonte de cinco años, con liquidación al quinto año y sin perpetuidad; el flujo de caja libre es positivo desde el primer año. Estos resultados son proyecciones condicionadas a la validación comercial del piloto, en particular del precio de S/ 140; el detalle se presenta en el capítulo VI.")
    S('Cada paquete deja un margen de contribución de S/ 39.13 sin', f"Cada paquete deja un margen de contribución de S/ {n2(UNI['B14'].value)} sin IGV ({pc(UNI['C14'].value)} del valor de venta). La frecuencia de seis paquetes por cliente al año se sustenta en el consumo de un hogar: un hogar limeño tiene 3.4 personas en promedio (Cámara de Comercio de Lima [CCL], s.f.) y el consumo per cápita bordea 1.6 a 1.7 kilogramos al año (El Peruano, 2026; Agraria.pe, 2025), es decir, entre 5.4 y 5.8 kilogramos por hogar. La preventa reduce la necesidad de capital de trabajo, porque el cliente paga antes de que la empresa compre el café al productor.")
    S('El primer año exige captar', f"El primer año exige captar 2,000 clientes, lo que con una conversión de 3% equivale a unas 66,667 visitas calificadas. La relación entre el valor del cliente y su costo de adquisición es de {lc:.2f} veces, por encima del mínimo de tres veces que suele exigirse en modelos de suscripción; sin embargo, depende de la retención y del CAC, que son los supuestos de menor confianza del modelo (Apéndice H).")
    S('Como el COK es', f"Como el COK es un supuesto relevante, la hoja de sensibilidad calcula el VAN con tasas de 12% a 28%: el VAN se mantiene positivo en todo ese rango, entre S/ {V(sens['28'])} con 28% y S/ {V(sens['12'])} con 12%. La tasa social de descuento del MEF, de 8%, se usa solo para el VAN social del capítulo VII.")
    set_par(find_par(doc, 'Nota. Valores sin IGV. El valor de venta'), 'Nota. Valores sin IGV. El valor de venta corresponde al precio de S/ 140 dividido entre 1.18.')
    S('Cada paquete deja un margen de contribución de S/ 39.13 (', f"Cada paquete deja un margen de contribución de S/ {n2(UNI['B14'].value)} ({pc(UNI['C14'].value)} del valor de venta). El café pagado al productor es el mayor costo y representa alrededor de 25% del valor de venta: una parte importante de lo que paga el cliente llega a la finca.")
    S('El resultado muestra una ventaja', "El resultado muestra una ventaja del modelo: desde el primer año, los anticipos de la preventa cubren el inventario de café y de empaques, y el capital de trabajo requerido es nulo; por eso no hay recupero de capital de trabajo al quinto año. Esta ventaja depende del supuesto de anticipos, de baja confianza, que debe confirmarse en el piloto.")
    S('La Tabla 23 presenta', f"La Tabla 23 presenta el estado de resultados. El primer año ya registra utilidad operativa positiva, aunque pequeña, porque absorbe la deducción de los gastos preoperativos y el costo de captar a los primeros clientes; el margen EBITDA pasa de {pc(EER['B22'].value)} en el primer año a {pc(m5)} en el quinto.")
    S('La Tabla 24 presenta', "La Tabla 24 presenta el flujo de caja libre. Solo el flujo del año 0 es negativo, por la inversión y los kits del primer año; desde el año 1 el flujo es positivo (S/ " + V(FCL['C17'].value) + "). En el quinto año el proyecto se liquida con el valor de rescate de S/ 154,500, sin perpetuidad y sin recupero de capital de trabajo.")
    S('La Tabla 25 resume', f"La Tabla 25 resume los indicadores del caso base. Con los supuestos actuales, el VAN es positivo al COK de 20.94%, la TIR supera ampliamente el COK y la inversión se recupera, en valor presente, en menos de dos años ({FCL['B26'].value:.2f} años).")
    S('El análisis costo-beneficio', f"El análisis costo-beneficio confirma la conveniencia del proyecto en el caso base: por cada sol de egresos en valor presente, incluidos la inversión, los costos y los impuestos, el modelo genera S/ {FCL['B29'].value:.2f} de ingresos en valor presente. La relación sigue siendo estrecha porque el café comprado al productor es un costo alto por diseño, y es sensible a la clasificación de los flujos y al valor de rescate; por eso se complementa con el VAN esperado por escenarios, de S/ {V(SEN['B17'].value)}.")
    S('Se calcularon dos puntos', f"Se calcularon dos puntos de equilibrio. El operativo divide los costos fijos y la inversión en adquisición de clientes, netos del margen de los tours, entre el margen unitario neto del marketing de retención; el contable añade la depreciación. En el primer año, el equilibrio operativo exige {V(pe1)} paquetes frente a 12,000 proyectados, con un margen de seguridad de {pc(ms1)}; desde el segundo año, las ventas superan el equilibrio con un margen de seguridad de {pc(ms2)} (Tabla 26).")
    S('El precio pagado por el cliente es', f"El precio pagado por el cliente sigue siendo la variable más sensible: el VAN se anula con una caída de 25.6%, es decir, con un precio de S/ 104, fuera del rango de ±20% de la Tabla 27. Con el precio de S/ 140 el VAN permanece positivo en todas las variaciones de ±20% de la tabla. La Tabla 28 combina varias variables en cuatro escenarios, incluido uno de estrés, con probabilidades asignadas por el equipo.")
    S('Los escenarios pesimista', f"Los cuatro escenarios dan VAN positivo, incluido el de estrés (S/ {V(SEN['B16'].value)}), que queda cerca de cero. El VAN esperado es de S/ {V(SEN['B17'].value)}. Estos escenarios mueven el precio solo entre 0% y -10% de S/ 140 (hasta S/ 126), por lo que no cubren el caso en que el consumidor no acepte pagar S/ 140. La conclusión es clara: el valor del proyecto depende de sostener el precio y el volumen de la preventa, y no debe invertirse a escala antes de validarlos.")
    S('El VANS asciende', f"El VANS asciende a S/ {V(vans)}. Alrededor de un tercio corresponde al ingreso incremental de los productores, lo que muestra que una parte importante del valor creado queda en las familias productoras. Como los supuestos del productor provienen de referencias regionales, el VANS debe recalcularse con los rendimientos y costos medidos en las fincas piloto.")
    S('Cuarto, con los', f"Cuarto, con los supuestos actuales el modelo es viable: el flujo de caja libre es positivo desde el primer año y, con un COK de 20.94%, el VAN es de S/ {V(base)}, la TIR de {pc(FCL['B24'].value)} y la relación beneficio-costo de {FCL['B29'].value:.2f}; el VAN esperado por escenarios es de S/ {V(SEN['B17'].value)}. Una simulación Monte Carlo de 10,000 iteraciones da una mediana de S/ {MCV['mediana']} y una probabilidad de VAN negativo de {MCV['p']}. El precio sigue siendo la variable crítica: el VAN se anula con un precio de S/ 104, y el precio de S/ 140 aún no se ha probado con compra real. Estos resultados son proyecciones condicionadas a la validación comercial.")
    S('Quinto, el modelo', f"Quinto, el modelo es socialmente rentable con los supuestos actuales: el ingreso neto estimado del productor sube en 64.7% con la tecnificación madura y el VAN social al 8% alcanza S/ {V(vans)}. Además, el modelo no depende del Estado en ninguna de sus etapas.")
    S('Se recomienda ejecutar', "Se recomienda ejecutar el piloto antes de la inversión completa y no escalar si la preventa de prueba no alcanza 3% de conversión y 300 paquetes, o si menos de la mitad de los productores invitados firma el acuerdo. El piloto debe probar el precio de S/ 140 frente a S/ 130 y S/ 120, y medir el costo real por visita calificada, porque el presupuesto del piloto (S/ 0.60 por visita) es cuatro veces menor que el del modelo (S/ 2.40). Con los resultados, se deben actualizar los supuestos del modelo financiero, las probabilidades de los escenarios, la simulación Monte Carlo y el VAN social.")
    # nota de la Tabla 32
    for p in doc.paragraphs:
        if 'unos S/ 30,500 sin IGV' in p.text:
            for r in p.runs: r.text = r.text.replace('unos S/ 30,500 sin IGV', 'unos S/ 35,600 sin IGV, a S/ 118.64 por paquete')
