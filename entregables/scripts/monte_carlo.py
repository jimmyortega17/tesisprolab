"""Simulación Monte Carlo del VAN, la TIR y el LTV de Cosecha Viva.
Uso: python3 monte_carlo.py   (semilla fija; 10,000 iteraciones)
Cada variable de baja confianza (Apéndice H) se muestrea con una distribución triangular (mínimo, más probable, máximo)."""
import numpy as np, json
from modelo import modelo, BASE

N = 10000
rng = np.random.default_rng(2026)
COSTO_VISITA = BASE['cac'] * 0.03          # S/ 2.40 por visita calificada (CAC 80 con conversión 3%)

# nombre, etiqueta, (mín, moda, máx), justificación
DIST = [
    ('precio',  'Precio del paquete (S/, con IGV)',          (100, 120, 135), 'Quiebre del VAN en S/ 104.40; referencia de mercado cercana a S/ 130 por kg (Mercado Libre, 2026)'),
    ('paqf',    'Factor de paquetes vendidos',               (0.70, 1.00, 1.15), 'Escenarios del equipo: estrés -30%, pesimista -20%, optimista +15%'),
    ('conv',    'Conversión de visitas en compra',           (0.015, 0.03, 0.04), 'Umbral H5 de 3%; CAC = costo por visita / conversión'),
    ('ret',     'Retención anual de clientes',               (0.40, 0.60, 0.75), 'Confianza baja; validar en el piloto (H8)'),
    ('ppc',     'Paquetes por cliente al año',               (4, 6, 7), 'Hogar limeño de 3.4 personas; validar recompra (H8)'),
    ('p_emp',   'Precio pagado por el café pergamino (S/ por kg)', (18, 20, 25), 'Alza del precio en chacra; el compromiso es una prima sobre el mercado'),
    ('r1',      'Rendimiento con tecnificación (qq/ha)',     (18, 25, 30), 'H4: umbral de 22 qq/ha; productores con nutrición balanceada superan 50 qq/ha'),
    ('sel',     'Proporción de café selecto',                (0.30, 0.40, 0.45), 'Supuesto del equipo; medir en fincas piloto'),
    ('f_cf',    'Factor de costos fijos',                    (0.95, 1.00, 1.25), 'Sobrecosto operativo más probable que ahorro'),
    ('f_var',   'Factor de maquila, empaque y envío',        (0.90, 1.00, 1.20), 'Cotizaciones no cerradas con tostaduría ni mensajería'),
    ('visitan', 'Clientes activos que visitan la finca',     (0.02, 0.05, 0.07), 'Umbral H8 de 5%'),
]

def draw():
    return {k: rng.triangular(*r, size=N) for k, _, r, _ in DIST}

def run(sample, i):
    s = {k: v[i] for k, v in sample.items()}
    base = BASE
    kw = dict(precio=s['precio'], paq=[x * s['paqf'] for x in base['paq']], cac=COSTO_VISITA / s['conv'],
              ret=s['ret'], ppc=s['ppc'], p_emp=s['p_emp'], r1=s['r1'], sel=s['sel'], f_cf=s['f_cf'],
              maquila=base['maquila'] * s['f_var'], caja=base['caja'] * s['f_var'], envio=base['envio'] * s['f_var'],
              visitan=s['visitan'])
    return modelo(**kw)

if __name__ == '__main__':
    smp = draw()
    out = [run(smp, i) for i in range(N)]
    van = np.array([o['van'] for o in out]); ltv = np.array([o['ltv'] for o in out])
    cac = np.array([o['cac'] for o in out]); lc = ltv / cac
    tir = np.array([o['tir'] for o in out])
    def st(x):
        return dict(media=float(np.nanmean(x)), mediana=float(np.nanmedian(x)), desv=float(np.nanstd(x)),
                    p5=float(np.nanpercentile(x, 5)), p25=float(np.nanpercentile(x, 25)),
                    p75=float(np.nanpercentile(x, 75)), p95=float(np.nanpercentile(x, 95)),
                    min=float(np.nanmin(x)), max=float(np.nanmax(x)))
    # correlación de rangos (Spearman) con el VAN
    from scipy.stats import spearmanr
    corr = {k: float(spearmanr(smp[k], van)[0]) for k, *_ in DIST}
    corr_ltv = {k: float(spearmanr(smp[k], ltv)[0]) for k, *_ in DIST}
    res = dict(n=N, van=st(van), ltv=st(ltv), ltv_cac=st(lc), tir=st(tir[~np.isnan(tir)]),
               p_van_neg=float((van < 0).mean()), p_lc_menor3=float((lc < 3).mean()),
               p_van_mayor_cero_y_tir=float((tir > out[0]['fcl'][0] * 0 + 0.2094).mean()),
               corr_van=corr, corr_ltv=corr_ltv,
               van_array=van.tolist(), ltv_array=ltv.tolist(), lc_array=lc.tolist(),
               dist=[(k, l, r, j) for k, l, r, j in DIST])
    json.dump(res, open('mc_resultados.json', 'w'))
    print(json.dumps({k: v for k, v in res.items() if k not in ('van_array', 'ltv_array', 'lc_array', 'dist')}, indent=1))
