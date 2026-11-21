import json, numpy as np, openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as L
from modelo import modelo

R = json.load(open('mc_resultados.json'))
wb = openpyxl.load_workbook('original.xlsx')
wb['Supuestos']['B6'] = 140
wb['Supuestos']['D6'] = 'Decisión del equipo: S/ 140 para que el flujo del año 1 no sea negativo; referencia de café de especialidad de Chanchamayo a S/ 130 en Mercado Libre Perú (2026)'
F = 'Times New Roman'
NUM = '#,##0;(#,##0);-'; PCT = '0.0%;(0.0%);-'; DEC = '#,##0.00'
bold = Font(name=F, bold=True, size=11); norm = Font(name=F, size=11); inp = Font(name=F, size=11, color='0000FF')
hdr_fill = PatternFill('solid', fgColor='E8EFE0'); thin = Side(style='thin', color='999999')

def put(ws, ref, v, f=norm, nf=None, fill=None, wrap=False):
    c = ws[ref]; c.value = v; c.font = f
    if nf: c.number_format = nf
    if fill: c.fill = fill
    if wrap: c.alignment = Alignment(wrap_text=True, vertical='top')
    return c
def head(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = put(ws, f'{L(col+i)}{row}', t, bold, fill=hdr_fill); c.alignment = Alignment(wrap_text=True, vertical='center')

# ---------- 0. VAN social = beneficios sociales - costos sociales (sin FCL ni inversión) ----------
FF = 'Times New Roman'
def sty(c, b=False, nf=None, fill=None):
    c.font = Font(name=FF, size=11, bold=b)
    if nf: c.number_format = nf
    if fill: c.fill = PatternFill('solid', fgColor=fill)
pr = wb['Productor']
pr.delete_rows(15, 20)
pr['A1'] = 'Beneficio del productor por cohortes'
pr['A2'] = 'Soles por año. Valoración a precios de mercado, sin precios sociales. Base de las hojas Beneficios Sociales, Costos Sociales y VAN Social.'
for r_, lab, f_ in ((15, 'Ingreso bruto anual del productor (S/)', '=X11*X7*X9+X11*(1-X7)*X8'),
                    (16, 'Beneficio bruto incremental frente a la situación sin el modelo (S/)', '=X15-$B15'),
                    (17, 'Costo adicional de insumos y cosecha selectiva (S/)', '=X10*X5')):
    pr.cell(r_, 1, lab); sty(pr.cell(r_, 1))
    for col in 'BCD':
        if r_ == 16 and col == 'B': continue
        c = pr[f'{col}{r_}']; c.value = f_.replace('X', col); sty(c, nf=NUM)
pr['A18'] = 'Comprobación: ingreso neto incremental = beneficio bruto - costo adicional'; sty(pr['A18'])
pr['C18'] = '=C16-C17-C13'; pr['D18'] = '=D16-D17-D13'
for k in ('C18', 'D18'): sty(pr[k], nf=NUM)

yrs = ['Año 1', 'Año 2', 'Año 3', 'Año 4', 'Año 5']; dcols = 'CDEFG'
def sheet_head(ws, title, sub):
    ws.column_dimensions['A'].width = 66
    for c_ in 'BCDEFG': ws.column_dimensions[c_].width = 16
    ws['A1'] = title; sty(ws['A1'], b=True); ws['A2'] = sub; sty(ws['A2'])
    ws['A4'] = 'Criterio'; sty(ws['A4'], b=True, fill='E8EFE0')
    for i, y in enumerate(yrs):
        c = ws.cell(4, 2 + i, y); sty(c, b=True, fill='E8EFE0')

bs = wb.create_sheet('Beneficios Sociales')
sheet_head(bs, 'Beneficios sociales', 'Un beneficio social por cada ODS declarado en la tesis (ODS 1, 2 y 8). Soles por año.')
rows_b = [(5, 'DATOS BASE', None, True),
          (6, 'Productores con tecnificación madura', '=Demanda!{d}17', False),
          (7, 'Productores en su primera campaña', '=Demanda!{d}18', False),
          (8, 'ODS 2 HAMBRE CERO (META 2.3): VALOR DE LA PRODUCCIÓN ADICIONAL', None, True),
          (9, 'Producción adicional por productor maduro (qq por año)', '=Productor!$C$11-Productor!$B$11', False),
          (10, 'Precio del intermediario (S/ por qq)', '=Productor!$B$8', False),
          (11, 'Beneficio social 1 (ODS 2: más producción por mayor productividad)', '={c}6*{c}9*{c}10+{c}7*(Productor!$D$11-Productor!$B$11)*{c}10', True),
          (12, 'ODS 1 FIN DE LA POBREZA (META 1.2): SOBREPRECIO DEL CAFÉ SELECTO', None, True),
          (13, 'Café selecto por productor maduro (qq por año)', '=Productor!$C$11*Productor!$C$7', False),
          (14, 'Café selecto por productor en su primera campaña (qq por año)', '=Productor!$D$11*Productor!$D$7', False),
          (15, 'Sobreprecio por qq frente al intermediario (S/)', '=Productor!$C$9-Productor!$C$8', False),
          (16, 'Beneficio social 2 (ODS 1: mayor ingreso por el sobreprecio)', '=({c}6*{c}13+{c}7*{c}14)*{c}15', True),
          (17, 'ODS 8 TRABAJO DECENTE Y CRECIMIENTO ECONÓMICO (METAS 8.3 Y 8.9): EMPLEO E INGRESO LOCAL', None, True),
          (18, 'Técnicos de campo contratados', '=Demanda!{d}23', False),
          (19, 'Remuneración anual por técnico, con cargas laborales (S/)', '=CostosFijos!$B$10*12*(1+Supuestos!$B$48)', False),
          (20, 'Beneficio social 3 (ODS 8: empleo calificado en el campo)', '={c}18*{c}19', True),
          (21, 'Gasto local de los tours: transporte, guía, almuerzo y pago al productor (S/)', '=-EERR!{c}9', False),
          (22, 'Beneficio social 4 (ODS 8: turismo en la finca)', '={c}21', True),
          (23, 'TOTAL DE BENEFICIOS SOCIALES', '={c}11+{c}16+{c}20+{c}22', True)]
for r_, lab, f_, b_ in rows_b:
    bs.cell(r_, 1, lab); sty(bs.cell(r_, 1), b=b_)
    if f_:
        for i in range(5):
            c = bs.cell(r_, 2 + i, f_.format(d=dcols[i], c=L(2 + i))); sty(c, b=b_, nf=NUM)
cs = wb.create_sheet('Costos Sociales')
sheet_head(cs, 'Costos sociales', 'Estimación del flujo de costos sociales: recursos que usa el productor para producir más y mejor café. Soles por año.')
rows_c = [(5, 'DATOS BASE', None, True),
          (6, 'Productores con tecnificación madura', "='Beneficios Sociales'!{c}6", False),
          (7, 'Productores en su primera campaña', "='Beneficios Sociales'!{c}7", False),
          (8, 'A) COSTO ADICIONAL DE INSUMOS Y COSECHA SELECTIVA, PRODUCTORES MADUROS', None, True),
          (9, 'Costo adicional por productor (S/ por año)', '=Productor!$C$17', False),
          (10, 'Costo social 1 (productores maduros)', '={c}6*{c}9', False),
          (11, 'B) COSTO ADICIONAL EN LA PRIMERA CAMPAÑA (PLAN GRADUAL)', None, True),
          (12, 'Costo adicional por productor (S/ por año)', '=Productor!$D$17', False),
          (13, 'Costo social 2 (productores en primera campaña)', '={c}7*{c}12', False),
          (14, 'TOTAL DE COSTOS SOCIALES', '={c}10+{c}13', True)]
for r_, lab, f_, b_ in rows_c:
    cs.cell(r_, 1, lab); sty(cs.cell(r_, 1), b=b_)
    if f_:
        for i in range(5):
            c = cs.cell(r_, 2 + i, f_.format(c=L(2 + i))); sty(c, b=b_, nf=NUM)
vs = wb.create_sheet('VAN Social')
sheet_head(vs, 'VAN social', 'VAN social = beneficios sociales - costos sociales, descontados a la tasa social del MEF. No incluye el flujo de caja libre ni la inversión de la empresa.')
vrows = [(5, 'Beneficios sociales totales', "='Beneficios Sociales'!{c}23"), (6, 'Costos sociales', "='Costos Sociales'!{c}14"), (7, 'Flujo social (beneficios - costos)', '={c}5-{c}6')]
for r_, lab, f_ in vrows:
    vs.cell(r_, 1, lab); sty(vs.cell(r_, 1), b=(r_ == 7))
    for i in range(5):
        c = vs.cell(r_, 2 + i, f_.format(c=L(2 + i))); sty(c, b=(r_ == 7), nf=NUM)
vs['A9'] = 'Tasa de descuento social (MEF)'; sty(vs['A9']); vs['B9'] = '=COK!$B$13'; sty(vs['B9'], nf=PCT)
for i in range(5):
    vs.cell(11 + i, 1, f'Valor presente (VP) año {i + 1}'); sty(vs.cell(11 + i, 1))
    c = vs.cell(11 + i, 2, f'={L(2 + i)}7/(1+$B$9)^{i + 1}'); sty(c, nf=NUM)
vs['A17'] = 'VAN SOCIAL (S/)'; sty(vs['A17'], b=True, fill='E8EFE0'); vs['B17'] = '=SUM(B11:B15)'; sty(vs['B17'], b=True, nf=NUM, fill='E8EFE0')
vs['A18'] = 'VP de los beneficios sociales'; vs['B18'] = '=NPV($B$9,B5:F5)'
vs['A19'] = 'VP de los costos sociales'; vs['B19'] = '=NPV($B$9,B6:F6)'
vs['A20'] = 'Relación beneficio-costo social'; vs['B20'] = '=B18/B19'
for r_ in (18, 19): sty(vs.cell(r_, 1)); sty(vs.cell(r_, 2), nf=NUM)
sty(vs['A20']); sty(vs['B20'], nf='0.00')
vs['A22'] = 'Nota. No hay año 0: el VAN social no incluye inversión ni flujo de caja libre. Los costos sociales son los recursos que el productor destina a insumos y cosecha selectiva; la prima pagada por la empresa es una transferencia y no se cuenta.'
sty(vs['A22'])
# quitar el VAN financiero a la tasa social de la hoja FCL
fcl_ws = wb['FCL']; fcl_ws['A30'] = None; fcl_ws['B30'] = None
for nm in ('VAN Social', 'Costos Sociales', 'Beneficios Sociales'):
    wb.move_sheet(nm, offset=wb.sheetnames.index('Productor') + 1 - wb.sheetnames.index(nm))

# ---------- FCL al WACC y FCA al COK ----------
fw = wb.create_sheet('FCA_WACC')
fw.column_dimensions['A'].width = 70
for c_ in 'BCDEFG': fw.column_dimensions[c_].width = 16
def put2(ref, v, b=False, nf=None, color=None, fill=None):
    c = fw[ref]; c.value = v; c.font = Font(name='Times New Roman', size=11, bold=b, color=color)
    if nf: c.number_format = nf
    if fill: c.fill = PatternFill('solid', fgColor=fill)
put2('A1', 'FCL descontado al WACC y FCA descontado al COK', True)
put2('A2', 'El FCL (flujo de caja libre) se descuenta con el WACC. El FCA (flujo de caja del accionista) se descuenta con el COK. Sin préstamo, el FCA es igual al FCL y el WACC es igual al COK.')
put2('A4', 'Financiamiento', True, fill='E8EFE0'); put2('B4', 'Valor', True, fill='E8EFE0')
put2('A5', 'Préstamo (S/): 0 = el proyecto se financia solo con aporte de los socios'); put2('B5', 0, color='0000FF', nf=NUM)
put2('A6', 'Tasa de interés anual del préstamo (referencial; no se usa si no hay préstamo)'); put2('B6', 0.12, color='0000FF', nf=PCT)
put2('A7', 'Plazo del préstamo (años, amortización constante)'); put2('B7', 5, color='0000FF')
put2('A8', 'Impuesto a la renta'); put2('B8', '=Supuestos!$B$47', nf=PCT)
put2('A9', 'Aporte de los socios (patrimonio) = salida de caja del año 0 - préstamo'); put2('B9', '=-FCL!$B$17-$B$5', nf=NUM)
put2('A10', 'Proporción de deuda D / (D + E)'); put2('B10', '=$B$5/($B$5+$B$9)', nf=PCT)
put2('A11', 'Proporción de patrimonio E / (D + E)'); put2('B11', '=$B$9/($B$5+$B$9)', nf=PCT)
put2('A12', 'COK (costo de oportunidad del accionista)'); put2('B12', '=COK!$B$12', nf=PCT)
put2('A13', 'Costo de la deuda después de impuestos = Kd x (1 - t)'); put2('B13', '=$B$6*(1-$B$8)', nf=PCT)
put2('A14', 'WACC = E/(D+E) x COK + D/(D+E) x Kd x (1 - t)', True); put2('B14', '=$B$11*$B$12+$B$10*$B$13', True, nf='0.00%')
for i, y in enumerate(['Año 0'] + yrs): put2(f'{L(2 + i)}16', y, True, fill='E8EFE0')
put2('A16', 'Flujos (S/)', True, fill='E8EFE0')
labs = {17: 'FCL (flujo de caja libre)', 18: '(+) Préstamo recibido', 19: 'Saldo del préstamo al cierre', 20: '(-) Amortización', 21: '(-) Intereses', 22: '(+) Escudo fiscal de los intereses', 23: 'FCA (flujo de caja del accionista)', 24: 'FCA - FCL'}
for r_, t_ in labs.items(): put2(f'A{r_}', t_, r_ in (17, 23))
for i in range(6):
    c = L(2 + i); pc_ = L(1 + i)
    put2(f'{c}17', f'=FCL!{c}17', True, NUM)
    put2(f'{c}18', '=$B$5' if i == 0 else 0, nf=NUM)
    put2(f'{c}19', '=B18' if i == 0 else f'={pc_}19-{c}20', nf=NUM)
    put2(f'{c}20', 0 if i == 0 else f'=MIN({pc_}19,$B$5/$B$7)', nf=NUM)
    put2(f'{c}21', 0 if i == 0 else f'={pc_}19*$B$6', nf=NUM)
    put2(f'{c}22', 0 if i == 0 else f'={c}21*$B$8', nf=NUM)
    put2(f'{c}23', f'={c}17+{c}18-{c}20-{c}21+{c}22', True, NUM)
    put2(f'{c}24', f'={c}23-{c}17', nf=NUM)
put2('A26', 'Resultados', True, fill='E8EFE0'); put2('B26', 'Valor', True, fill='E8EFE0')
put2('A27', 'VAN económico: FCL descontado al WACC (S/)', True); put2('B27', '=B17+NPV($B$14,C17:G17)', True, NUM)
put2('A28', 'VAN financiero: FCA descontado al COK (S/)', True); put2('B28', '=B23+NPV($B$12,C23:G23)', True, NUM)
put2('A29', 'Diferencia entre ambos VAN (0 si no hay préstamo)'); put2('B29', '=B27-B28', nf=NUM)
put2('A30', 'TIR económica (FCL)'); put2('B30', '=IRR(B17:G17)', nf=PCT)
put2('A31', 'TIR financiera (FCA)'); put2('B31', '=IRR(B23:G23)', nf=PCT)
put2('A32', 'TIR modificada del FCA (financiamiento y reinversión al COK)'); put2('B32', '=MIRR(B23:G23,$B$12,$B$12)', nf=PCT)
put2('A34', 'Nota. Si se ingresa un préstamo en B5, el FCA cambia (préstamo, amortización, intereses y escudo fiscal) y el WACC baja por la deuda. Este modelo no recalcula la beta ni el COK por mayor apalancamiento, y la hoja Sensibilidad sigue usando el COK.')
# la hoja FCL descuenta con el WACC
fl = wb['FCL']
for i in range(6): fl.cell(18, 2 + i).value = f'=1/(1+FCA_WACC!$B$14)^{i}'
fl['A18'] = 'Factor de descuento al WACC'
fl['A23'] = 'VAN económico: FCL descontado al WACC (S/)'
fl['A24'] = 'TIR económica (FCL)'
fl['A25'] = 'WACC (igual al COK si no hay préstamo)'; fl['B25'] = '=FCA_WACC!$B$14'
fl['A30'] = 'VAN financiero: FCA descontado al COK (S/)'; fl['B30'] = '=FCA_WACC!$B$28'; fl['B30'].number_format = NUM; fl['A30'].font = Font(name='Times New Roman', size=11)
fl['A30'].font = Font(name='Times New Roman', size=11); fl['B30'].font = Font(name='Times New Roman', size=11)
wb.move_sheet('FCA_WACC', offset=wb.sheetnames.index('FCL') + 1 - wb.sheetnames.index('FCA_WACC'))

# ---------- 1. recrear el gráfico de la hoja Sensibilidad (openpyxl no conserva gráficos al cargar) ----------
s = wb['Sensibilidad']
ch = LineChart(); ch.title = 'Sensibilidad del VAN (S/)'; ch.y_axis.title = 'VAN (S/)'; ch.x_axis.title = 'Variación'
for r in range(5, 10):
    ch.add_data(Reference(s, min_col=1, max_col=6, min_row=r, max_row=r), from_rows=True, titles_from_data=True)
ch.set_categories(Reference(s, min_col=2, max_col=6, min_row=4, max_row=4))
ch.width, ch.height = 18, 8; s.add_chart(ch, 'I4')

# ---------- 2. MC_Datos ----------
d = wb.create_sheet('MC_Datos')
put(d, 'A1', 'Resultados de las 10,000 iteraciones de la simulación Monte Carlo', bold)
put(d, 'A2', 'Valores generados con entregables/scripts/monte_carlo.py (semilla 2026). No se recalculan solos: si cambia un supuesto, vuelva a ejecutar el script.')
head(d, 4, ['Iteración', 'VAN (S/)', 'LTV (S/)', 'CAC (S/)', 'LTV / CAC'])
van = R['van_array']; ltv = R['ltv_array']; lc = R['lc_array']
for i in range(len(van)):
    r = 5 + i
    d.cell(r, 1, i + 1); d.cell(r, 2, round(van[i], 2)); d.cell(r, 3, round(ltv[i], 2))
    d.cell(r, 4, round(ltv[i] / lc[i], 2)); d.cell(r, 5, round(lc[i], 3))
    d.cell(r, 2).number_format = NUM; d.cell(r, 3).number_format = DEC; d.cell(r, 4).number_format = DEC; d.cell(r, 5).number_format = DEC
for col, w in zip('ABCDE', (12, 16, 14, 12, 12)): d.column_dimensions[col].width = w
d.freeze_panes = 'A5'
N = len(van); last = 4 + N
rngV = f'MC_Datos!$B$5:$B${last}'; rngL = f'MC_Datos!$C$5:$C${last}'; rngC = f'MC_Datos!$E$5:$E${last}'

# ---------- 3. MonteCarlo (resumen vivo sobre MC_Datos) ----------
m = wb.create_sheet('MonteCarlo')
m.column_dimensions['A'].width = 52
for c in 'BCDEFG': m.column_dimensions[c].width = 16
m.column_dimensions['E'].width = 70
put(m, 'A1', 'Simulación Monte Carlo del VAN y del LTV (10,000 iteraciones)', bold)
put(m, 'A2', 'Distribuciones triangulares (mínimo, más probable, máximo) para los supuestos de confianza media y baja del Apéndice H. COK fijo en 20.94%. Los estadísticos se calculan con fórmulas sobre la hoja MC_Datos.')
put(m, 'A4', 'A. Distribuciones de entrada', bold)
head(m, 5, ['Variable', 'Mínimo', 'Más probable', 'Máximo', 'Justificación'])
for i, (k, lab, (lo, mo, hi), jus) in enumerate(R['dist']):
    r = 6 + i
    put(m, f'A{r}', lab); 
    for col, v in zip('BCD', (lo, mo, hi)): put(m, f'{col}{r}', v, inp, '0.0%' if v < 1 and k not in ('paqf',) else '#,##0.00')
    put(m, f'E{r}', jus, wrap=True)
r0 = 6 + len(R['dist']) + 1
put(m, f'A{r0}', 'Nota. El CAC se simula como costo por visita calificada (S/ 2.40 = S/ 80 x 3%) dividido entre la conversión simulada.')
r1 = r0 + 2
put(m, f'A{r1}', 'B. Resultados', bold)
head(m, r1 + 1, ['Estadístico', 'VAN (S/)', 'LTV (S/)', 'LTV / CAC'])
stats = [('Media', 'AVERAGE({})'), ('Mediana', 'MEDIAN({})'), ('Desviación estándar', 'STDEV({})'), ('Percentil 5', 'PERCENTILE({},0.05)'),
         ('Percentil 25', 'PERCENTILE({},0.25)'), ('Percentil 75', 'PERCENTILE({},0.75)'), ('Percentil 95', 'PERCENTILE({},0.95)'),
         ('Mínimo', 'MIN({})'), ('Máximo', 'MAX({})')]
for i, (lab, f) in enumerate(stats):
    r = r1 + 2 + i; put(m, f'A{r}', lab)
    put(m, f'B{r}', '=' + f.format(rngV), nf=NUM); put(m, f'C{r}', '=' + f.format(rngL), nf=DEC); put(m, f'D{r}', '=' + f.format(rngC), nf='0.00')
r2 = r1 + 2 + len(stats) + 1
put(m, f'A{r2}', 'Probabilidad de VAN negativo'); put(m, f'B{r2}', f'=COUNTIF({rngV},"<0")/COUNT({rngV})', nf=PCT)
put(m, f'A{r2+1}', 'Probabilidad de LTV / CAC menor que 3'); put(m, f'B{r2+1}', f'=COUNTIF({rngC},"<3")/COUNT({rngC})', nf=PCT)
put(m, f'A{r2+2}', 'Probabilidad de VAN mayor que S/ 500,000'); put(m, f'B{r2+2}', f'=COUNTIF({rngV},">500000")/COUNT({rngV})', nf=PCT)
put(m, f'A{r2+3}', 'VAN del caso base (hoja FCL) y su percentil en la simulación'); put(m, f'B{r2+3}', '=FCL!$B$23', nf=NUM)
put(m, f'C{r2+3}', f'=COUNTIF({rngV},"<"&B{r2+3})/COUNT({rngV})', nf=PCT)
put(m, f'A{r2+4}', 'Nota. El caso base queda por encima de la mediana simulada: es un resultado optimista frente a la distribución completa.')

# histogramas con COUNTIFS vivos
def histo(title, rng, start_row, lo, hi, step, col0, nf):
    put(m, f'{L(col0)}{start_row}', title, bold)
    head(m, start_row + 1, ['Desde', 'Hasta', 'Frecuencia'], col0)
    edges = np.arange(lo, hi + step, step)
    for i in range(len(edges) - 1):
        r = start_row + 2 + i
        put(m, f'{L(col0)}{r}', float(edges[i]), nf=nf); put(m, f'{L(col0+1)}{r}', float(edges[i + 1]), nf=nf)
        a, b = f'{L(col0)}{r}', f'{L(col0+1)}{r}'
        put(m, f'{L(col0+2)}{r}', f'=COUNTIFS({rng},">="&{a},{rng},"<"&{b})', nf=NUM)
    n = len(edges) - 1
    bc = BarChart(); bc.type = 'col'; bc.title = title; bc.gapWidth = 20; bc.legend = None
    bc.add_data(Reference(m, min_col=col0 + 2, min_row=start_row + 1, max_row=start_row + 1 + n), titles_from_data=True)
    bc.set_categories(Reference(m, min_col=col0, min_row=start_row + 2, max_row=start_row + 1 + n))
    bc.y_axis.title = 'Frecuencia'; bc.width, bc.height = 17, 8
    return bc, start_row + 2 + n
h0 = r2 + 7
bc1, e1 = histo('Histograma del VAN (S/)', rngV, h0, -1800000, 2400000, 150000, 1, NUM)
m.add_chart(bc1, f'E{h0}')
h1 = e1 + 20
bc2, e2 = histo('Histograma del LTV (S/ por cliente)', rngL, h1, 150, 1100, 50, 1, '#,##0')
m.add_chart(bc2, f'E{h1}')
h2 = e2 + 20
bc3, e3 = histo('Histograma de LTV / CAC (veces)', rngC, h2, 1, 16, 1, 1, '0.0')
m.add_chart(bc3, f'E{h2}')
h3 = e3 + 3
put(m, f'A{h3}', 'C. Correlación de rangos (Spearman) con el VAN y el LTV', bold)
head(m, h3 + 1, ['Variable', 'Con el VAN', 'Con el LTV'])
lab = {k: l for k, l, _, _ in R['dist']}
for i, k in enumerate(sorted(R['corr_van'], key=lambda k: -abs(R['corr_van'][k]))):
    r = h3 + 2 + i; put(m, f'A{r}', lab[k]); put(m, f'B{r}', R['corr_van'][k], nf='0.00'); put(m, f'C{r}', R['corr_ltv'][k], nf='0.00')

# ---------- 4. Indicadores_Extra (todo con fórmulas) ----------
x = wb.create_sheet('Indicadores_Extra')
x.column_dimensions['A'].width = 62
for c in 'BCDEFG': x.column_dimensions[c].width = 15
put(x, 'A1', 'Indicadores adicionales: recupero simple, TIRM, necesidad de financiamiento y tornado', bold)
put(x, 'A2', 'Todas las celdas son fórmulas enlazadas a las hojas FCL, COK y Sensibilidad.')
head(x, 4, ['Concepto', 'Año 0', 'Año 1', 'Año 2', 'Año 3', 'Año 4', 'Año 5'])
put(x, 'A5', 'Flujo de caja libre (S/)')
put(x, 'A6', 'Flujo de caja libre acumulado (S/)')
for i in range(6):
    c = L(2 + i)
    put(x, f'{c}5', f'=FCL!{c}17', nf=NUM)
    put(x, f'{c}6', f'={c}5' if i == 0 else f'={L(1+i)}6+{c}5', nf=NUM)
put(x, 'A8', 'Periodo de recupero simple (años)')
put(x, 'B8', '=IF(G6<0,"Mayor a 5",IF(C6>=0,-B6/C5,IF(D6>=0,1-C6/D5,IF(E6>=0,2-D6/E5,IF(F6>=0,3-E6/F5,IF(G6>=0,4-F6/G5,"Mayor a 5"))))))', nf='0.00')
put(x, 'A9', 'Periodo de recupero descontado (años), hoja FCL'); put(x, 'B9', '=FCL!B26', nf='0.00')
put(x, 'A10', 'TIR (hoja FCL)'); put(x, 'B10', '=FCL!B24', nf=PCT)
put(x, 'A11', 'TIR modificada del FCA (financiamiento y reinversión al COK)'); put(x, 'B11', '=MIRR(B5:G5,COK!B12,COK!B12)', nf=PCT)
put(x, 'A12', 'COK'); put(x, 'B12', '=COK!B12', nf=PCT)
put(x, 'A13', 'Índice de rentabilidad (1 + VAN / VP de egresos)'); put(x, 'B13', '=1+FCL!B23/FCL!B28', nf='0.00')
put(x, 'A15', 'Necesidad de financiamiento', bold)
put(x, 'A16', 'Aporte inicial de los socios en el año 0 (inversión + kits + capital de trabajo)'); put(x, 'B16', '=-B5', nf=NUM)
put(x, 'A17', 'Necesidad máxima acumulada de caja (valle del flujo acumulado)'); put(x, 'B17', '=-MIN(B6:G6)', nf=NUM)
put(x, 'A18', 'Brecha que debe cubrirse además del aporte del año 0 (déficit del año 1)'); put(x, 'B18', '=MAX(0,B17-B16)', nf=NUM)
put(x, 'A19', 'Colchón recomendado (20% sobre la necesidad máxima, supuesto del equipo)'); put(x, 'B19', 0.2, inp, PCT)
put(x, 'A20', 'Aporte total recomendado de los socios (S/)'); put(x, 'B20', '=B17*(1+B19)', nf=NUM)
put(x, 'A21', 'Aporte por socio con 4 socios (S/)'); put(x, 'B21', '=B20/4', nf=NUM)
put(x, 'A22', 'Precio de quiebre del VAN (S/ con IGV), calculado con el script del modelo con las demás variables en su valor base'); put(x, 'B22', 104.22, inp, '0.00')
put(x, 'A23', 'Tornado: VAN con cada variable en -20% y +20% (enlazado a la hoja Sensibilidad)', bold)
head(x, 24, ['Variable', 'VAN con -20%', 'VAN con +20%', 'Amplitud (S/)', 'VAN base'])
for i, r_s in enumerate(range(5, 10)):
    r = 25 + i
    put(x, f'A{r}', f'=Sensibilidad!A{r_s}'); put(x, f'B{r}', f'=Sensibilidad!B{r_s}', nf=NUM); put(x, f'C{r}', f'=Sensibilidad!F{r_s}', nf=NUM)
    put(x, f'D{r}', f'=ABS(C{r}-B{r})', nf=NUM); put(x, f'E{r}', '=FCL!$B$23', nf=NUM)
tc = BarChart(); tc.type = 'bar'; tc.title = 'Tornado: VAN con variables en -20% y +20% (S/)'; tc.width, tc.height = 18, 8
tc.add_data(Reference(x, min_col=2, max_col=3, min_row=24, max_row=29), titles_from_data=True)
tc.set_categories(Reference(x, min_col=1, min_row=25, max_row=29)); x.add_chart(tc, 'G23')

# ---------- 5. Plan_Marketing (presupuesto con fórmulas) ----------
p = wb.create_sheet('Plan_Marketing')
p.column_dimensions['A'].width = 66
for c in 'BCDEFG': p.column_dimensions[c].width = 15
put(p, 'A1', 'Presupuesto del plan de marketing a cinco años', bold)
put(p, 'A2', 'Soles sin IGV. El total sale del modelo (hojas Demanda y EERR); la distribución por canal es un supuesto editable (celdas azules).')
head(p, 4, ['Concepto', 'Año 1', 'Año 2', 'Año 3', 'Año 4', 'Año 5'])
cols = 'CDEFG'; ecols = 'BCDEF'
put(p, 'A5', 'Clientes nuevos a captar'); put(p, 'A6', 'Visitas calificadas necesarias'); put(p, 'A7', 'Inversión en adquisición (clientes nuevos x CAC)')
put(p, 'A8', 'Marketing de retención y contenido (2% de ventas de café)'); put(p, 'A9', 'Total marketing (coincide con EERR)')
put(p, 'A10', 'Marketing / ventas netas'); put(p, 'A11', 'Costo implícito por visita calificada (CAC x conversión)')
for i in range(5):
    c = L(2 + i); dc = cols[i]; ec = ecols[i]
    put(p, f'{c}5', f'=Demanda!{dc}11', nf=NUM); put(p, f'{c}6', f'=Demanda!{dc}12', nf=NUM); put(p, f'{c}7', f'=Demanda!{dc}13', nf=NUM)
    put(p, f'{c}8', f'=EERR!{ec}5*Supuestos!$B$11', nf=NUM); put(p, f'{c}9', f'={c}7+{c}8', nf=NUM)
    put(p, f'{c}10', f'={c}9/EERR!{ec}7', nf=PCT); put(p, f'{c}11', '=Supuestos!$B$10*Supuestos!$B$12', nf=DEC)
put(p, 'A13', 'Distribución de la inversión en adquisición por canal (supuesto del equipo)', bold)
head(p, 14, ['Canal', '% del total', 'Año 1 (S/)', 'Año 3 (S/)', 'Año 5 (S/)'])
canales = [('Pauta digital (Meta y Google)', .45), ('Alianzas con tostadurías, cafeterías y hoteles (referidos)', .15),
           ('Marketing de contenidos y videos de cosecha (orgánico)', .15), ('Regalos corporativos y ventas B2B', .15), ('Programa de referidos entre clientes', .10)]
for i, (n, pc) in enumerate(canales):
    r = 15 + i; put(p, f'A{r}', n); put(p, f'B{r}', pc, inp, PCT)
    put(p, f'C{r}', f'=$B{r}*$B$7', nf=NUM); put(p, f'D{r}', f'=$B{r}*$D$7', nf=NUM); put(p, f'E{r}', f'=$B{r}*$F$7', nf=NUM)
put(p, 'A20', 'Total'); put(p, 'B20', '=SUM(B15:B19)', nf=PCT); put(p, 'C20', '=SUM(C15:C19)', nf=NUM); put(p, 'D20', '=SUM(D15:D19)', nf=NUM); put(p, 'E20', '=SUM(E15:E19)', nf=NUM)
put(p, 'A22', 'Verificación contra el piloto', bold)
put(p, 'A23', 'Pauta del piloto (Tabla 32, S/)'); put(p, 'B23', 6000, inp, NUM)
put(p, 'A24', 'Visitas calificadas que compraría esa pauta (según la tesis)'); put(p, 'B24', 10000, inp, NUM)
put(p, 'A25', 'Costo por visita implícito en el piloto (S/)'); put(p, 'B25', '=B23/B24', nf=DEC)
put(p, 'A26', 'Costo por visita implícito en el modelo (S/)'); put(p, 'B26', '=B11', nf=DEC)
put(p, 'A27', 'Diferencia (veces)'); put(p, 'B27', '=B26/B25', nf='0.0')
put(p, 'A28', 'Si el piloto paga la visita al costo del modelo, las 10,000 visitas cuestan (S/)'); put(p, 'B28', '=B24*B26', nf=NUM)
put(p, 'A29', 'Nota. El piloto presupuesta una visita mucho más barata que el modelo; el piloto debe medir el costo real por visita calificada.')

wb.move_sheet('MC_Datos', offset=0)
wb.calculation.fullCalcOnLoad = True
wb.save('../Modelo_Financiero_Tesis_Cosecha_Viva_v5.xlsx')
print('guardado')
