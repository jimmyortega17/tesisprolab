import openpyxl
wb = openpyxl.load_workbook('valores_v5.xlsx', data_only=True)
EER, FCL, CT, UNI, SEN, PRO, DEM = (wb[n] for n in ('EERR', 'FCL', 'CapitalTrabajo', 'Unitario', 'Sensibilidad', 'Productor', 'Demanda'))

def n0(v):
    v = round(v)
    return f'({abs(v):,})' if v < 0 else f'{v:,}'
def n2(v): return f'{v:,.2f}'
def pc(v, d=1): return f'{v*100:.{d}f}%'
def row(ws, r, c0, c1): return [ws.cell(r, c).value for c in range(c0, c1 + 1)]
import json
_R = json.load(open('mc_resultados.json'))
MCV = {'p': f"{_R['p_van_neg']*100:.1f}%", 'mediana': f"{_R['van']['mediana']:,.0f}"}
