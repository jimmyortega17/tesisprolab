"""Réplica en Python del modelo financiero de Cosecha Viva (hojas Supuestos a FCL del Excel v4).
Se valida contra los resultados del Excel (VAN base, sensibilidad y escenarios)."""
import math
import numpy as np

def xround(x):  # ROUND de Excel (mitad hacia arriba)
    return math.floor(x + 0.5)

def xroundup(x):
    return math.ceil(x - 1e-9)

BASE = dict(
    precio=140.0, igv=0.18, ppc=6.0, ret=0.60, cac=80.0, mkt=0.02, pasarela=0.042,
    p_inter=15.0, p_emp=20.0, verde=0.80, tostado=0.83, kgqq=46.0, pilado=0.8,
    maquila=8.0, bolsa=1.8, nbolsas=4, caja=2.5, envio=8.5,
    visitan=0.05, personas=2.5, tour_precio=180.0, tour_var=0.60,
    ha=2.0, r0=15.0, r1=25.0, sel=0.40, adic=2500.0, adic1=0.5, kit=3500.0,
    tec_por=30, inv_caf=4.0, inv_emp=1.0, caja_min=1.0, anticipo=1.5,
    tax=0.295, cargas=0.25, crec=0.05, rec_kit=0.5, rec_eq=1.0,
    paq=[12000, 26000, 40000, 54000, 66000], f_cf=1.0, cok=None,
)
COK = ((1 + 0.0488 + 3.56 * 0.0414 + 0.0118) * (1 + 0.0225) / (1 + 0.02133333333333333)) - 1
BASE['cok'] = COK

# inversión (con 10% de imprevistos): (monto, tasa de amortización, es_preoperativo)
INV = [(45000, 0.20, False), (22000, 0.25, False), (18000, 0.10, False), (9000, 0.10, False),
       (8000, 1.0, True), (30000, 1.0, True), (10000, 1.0, True)]
SALARIOS = 4500 + 4000 + 3000 + 2200 + 1500
OTROS = 800 + 1200 + 2000 + 1500 + 600
TEC = 2500

def irr(cf):
    lo, hi = -0.99, 10.0
    f = lambda r: sum(c / (1 + r) ** i for i, c in enumerate(cf))
    if f(lo) * f(hi) > 0:
        return float('nan')
    for _ in range(200):
        mid = (lo + hi) / 2
        if f(lo) * f(mid) <= 0: hi = mid
        else: lo = mid
    return (lo + hi) / 2

def modelo(**kw):
    p = dict(BASE); p.update(kw)
    pv = p['precio'] / (1 + p['igv'])                       # valor de venta
    kgp = 1 / (p['verde'] * p['tostado'])                   # kg pergamino por paquete
    compra = kgp * p['p_emp']; pil = kgp * p['pilado']
    emp = p['bolsa'] * p['nbolsas'] + p['caja']
    cvu = compra + pil + p['maquila'] + emp + p['envio'] + p['pasarela'] * p['precio']
    mc = pv - cvu
    tour_pv = p['tour_precio'] / (1 + p['igv'])
    u0 = p['ha'] * p['r0'] * p['sel']; u1 = p['ha'] * p['r1'] * p['sel']
    paq = p['paq']
    act = [xround(q / p['ppc']) for q in paq]
    ret = [0] + [xround(act[i - 1] * p['ret']) for i in range(1, 5)]
    nuevos = [max(0, a - r) for a, r in zip(act, ret)]
    cacc = [n * p['cac'] for n in nuevos]
    mad = []; new = []; tot = []; tec = []
    prev = 0
    for t in range(5):
        qq = paq[t] * kgp / p['kgqq']
        m = prev
        n = xroundup(max(0, qq - m * u1) / u0)
        mad.append(m); new.append(n); prev = m + n; tot.append(prev); tec.append(xroundup(prev / p['tec_por']))
    vis = [xround(a * p['visitan'] * p['personas']) for a in act]
    ventas_p = [q * pv for q in paq]; ventas_t = [v * tour_pv for v in vis]
    ventas = [a + b for a, b in zip(ventas_p, ventas_t)]
    cv_cafe = [q * cvu for q in paq]; cv_tour = [v * p['tour_var'] for v in ventas_t]
    mkt = [cacc[i] + ventas_p[i] * p['mkt'] for i in range(5)]
    cf = [(SALARIOS + tec[t] * TEC) * 12 * (1 + p['cargas']) * 1.0 + OTROS * 12 * (1 + p['crec']) ** t for t in range(5)]
    cf = [c * p['f_cf'] for c in cf]
    ebitda = [ventas[i] - cv_cafe[i] - cv_tour[i] - mkt[i] - cf[i] for i in range(5)]
    # inversión y depreciación
    imp = 1.10
    inv0 = sum(m * imp for m, _, _ in INV)
    fijos = sum(m * imp for m, _, pre in INV if not pre)
    kits = [new[t] * p['kit'] for t in range(5)] + [0]
    dep = [0] * 6; cum = [0.0] * len(INV)
    for y in range(1, 6):
        d = 0
        for k, (m, tasa, pre) in enumerate(INV):
            c = m * imp
            a = min(c * tasa, c - cum[k]); cum[k] += a; d += a
        d += sum(kits[:y]) * 0.1
        dep[y] = d
    dep_noprev = sum(cum[k] for k, (_, _, pre) in enumerate(INV) if not pre)
    vl_eq = fijos - dep_noprev
    vl_kits = sum(kits[:5]) - sum(sum(kits[:y]) * 0.1 for y in range(1, 6))
    rescate = vl_eq * p['rec_eq'] + vl_kits * p['rec_kit']
    ebit = [ebitda[i] - dep[i + 1] for i in range(5)]
    perd = 0.0; tax = []
    for i in range(5):
        renta = max(0, ebit[i] - perd); perd = max(0, perd - ebit[i]); tax.append(renta * p['tax'])
    ct = []
    for t in range(5):
        v = paq[t] * (compra + pil) * p['inv_caf'] / 12 + paq[t] * emp * p['inv_emp'] / 12 \
            + cf[t] * p['caja_min'] / 12 - ventas_p[t] * p['anticipo'] / 12
        ct.append(max(0, v))
    dct = [ct[0]] + [max(0, ct[t] - ct[t - 1]) for t in range(1, 5)] + [0]
    ing = [0] + ventas; ing[5] += sum(dct[:5]) + rescate
    egr = [inv0 + kits[0] + dct[0]]
    for y in range(1, 6):
        egr.append(kits[y] + cv_cafe[y - 1] + cv_tour[y - 1] + mkt[y - 1] + cf[y - 1] + tax[y - 1] + dct[y])
    fcl = [ing[i] - egr[i] for i in range(6)]
    van = sum(fcl[i] / (1 + p['cok']) ** i for i in range(6))
    ltv = p['ppc'] * (mc - p['mkt'] * pv) / (1 - p['ret'])
    return dict(van=van, tir=irr(fcl), fcl=fcl, ltv=ltv, cac=p['cac'], ltv_cac=ltv / p['cac'], mc=mc,
                productores=tot, ventas=ventas, ebitda=ebitda, nuevos=nuevos, act=act, cvu=cvu, pv=pv,
                inv0=inv0, ct=ct, kits=kits, tec=tec)

def escenario(precio=1, paq=1, cafe=1, cf=1, cac=1):
    return modelo(precio=BASE['precio'] * precio, paq=[x * paq for x in BASE['paq']], p_emp=20 * cafe, pilado=0.8 * cafe, f_cf=cf, cac=80 * cac)

if __name__ == '__main__':
    b = modelo()
    print('VAN base', round(b['van'], 2), 'TIR', round(b['tir'], 4), 'LTV', round(b['ltv'], 1))
    print('FCL', [round(x, 1) for x in b['fcl']])
    for name, kw in [('precio -20%', dict(precio=.8)), ('paquetes -20%', dict(paq=.8)), ('cafe +20%', dict(cafe=1.2)),
                     ('cf -20%', dict(cf=.8)), ('cac -20%', dict(cac=.8)),
                     ('pesimista', dict(precio=.95, paq=.8, cafe=1.2, cf=1.1, cac=1.2)),
                     ('optimista', dict(paq=1.15, cf=.95, cac=.9)),
                     ('estres', dict(precio=.9, paq=.7, cafe=1.2, cf=1.1, cac=1.3))]:
        print(name, round(escenario(**kw)['van'], 1))
