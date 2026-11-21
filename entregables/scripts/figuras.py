import json, numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt, matplotlib.ticker
from matplotlib.patches import FancyBboxPatch
from modelo import modelo, escenario, BASE
OUT = '../figuras/'
import os; os.makedirs(OUT, exist_ok=True)
R = json.load(open('mc_resultados.json'))
van = np.array(R['van_array']); ltv = np.array(R['ltv_array']); lc = np.array(R['lc_array'])
G, V, A, RED = '#4A6B3A', '#8B5A2B', '#C9A24B', '#B03A2E'
plt.rcParams.update({'font.size': 10, 'font.family': 'DejaVu Sans'})

def hist(x, titulo, xlabel, fname, ref=None, fmt='{:,.0f}', cut=None, color=G):
    fig, ax = plt.subplots(figsize=(7.2, 4))
    n, bins, patches = ax.hist(x, bins=40, color=color, edgecolor='white')
    if cut is not None:
        for b, p in zip(bins[:-1], patches):
            if b + (bins[1]-bins[0]) / 2 < cut: p.set_facecolor(RED)
        ax.axvline(cut, color='k', ls='--', lw=1)
    if ref is not None:
        ax.axvline(ref, color=A, lw=2); ax.text(ref, ax.get_ylim()[1]*0.95, ' caso base', color='#7a5d10', va='top')
    ax.axvline(np.median(x), color='#1f4e79', lw=1.5, ls=':'); 
    ax.text(np.median(x), ax.get_ylim()[1]*0.82, ' mediana ' + fmt.format(np.median(x)), color='#1f4e79')
    ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v,p: f'{v:,.0f}' if fmt=='{:,.0f}' else f'{v:,.1f}')); ax.set_title(titulo, fontweight='bold'); ax.set_xlabel(xlabel); ax.set_ylabel('Frecuencia (10,000 iteraciones)')
    ax.spines[['top', 'right']].set_visible(False)
    fig.tight_layout(); fig.savefig(OUT + fname, dpi=200); plt.close(fig)

b = modelo()
hist(van, 'Distribución del VAN (simulación Monte Carlo)', 'VAN al COK de 20.94% (S/)', 'fig_mc_van.png', ref=b['van'], cut=0)
hist(ltv, 'Distribución del LTV por cliente', 'LTV (S/ por cliente)', 'fig_mc_ltv.png', ref=b['ltv'], color=V)
hist(lc, 'Distribución de la relación LTV / CAC', 'LTV / CAC (veces)', 'fig_mc_ltvcac.png', ref=b['ltv']/80, fmt='{:,.1f}', cut=3, color='#5B7F95')

# tornado determinista (± 20%) con el modelo
casos = [('Precio del paquete', dict(precio=.8), dict(precio=1.2)), ('Paquetes vendidos', dict(paq=.8), dict(paq=1.2)),
         ('Precio del café pergamino', dict(cafe=1.2), dict(cafe=.8)), ('Costos fijos', dict(cf=1.2), dict(cf=.8)),
         ('CAC', dict(cac=1.2), dict(cac=.8))]
tor = []
for n, lo, hi in casos:
    tor.append((n, escenario(**lo)['van'], escenario(**hi)['van']))
tor.sort(key=lambda t: abs(t[2] - t[1]))
fig, ax = plt.subplots(figsize=(7.2, 3.8))
for i, (n, lo, hi) in enumerate(tor):
    ax.barh(i, lo - b['van'], left=b['van'], color=RED); ax.barh(i, hi - b['van'], left=b['van'], color=G)
ax.set_yticks(range(len(tor))); ax.set_yticklabels([t[0] for t in tor]); ax.axvline(b['van'], color='k', lw=1)
ax.axvline(0, color='grey', ls='--', lw=.8)
ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v,p: f'{v:,.0f}')); ax.set_xlabel('VAN (S/) con cada variable en ±20% del caso base'); ax.set_title('Tornado de sensibilidad del VAN', fontweight='bold')
ax.spines[['top', 'right']].set_visible(False)
import matplotlib.patches as mp
ax.legend(handles=[mp.Patch(color=RED, label='Efecto desfavorable'), mp.Patch(color=G, label='Efecto favorable')], loc='lower right', frameon=False)
fig.tight_layout(); fig.savefig(OUT + 'fig_tornado.png', dpi=200); plt.close(fig)
json.dump(tor, open('tornado.json', 'w'))

# tornado de correlación (Monte Carlo)
labels = {k: l for k, l, _, _ in R['dist']}
cv = sorted(R['corr_van'].items(), key=lambda t: abs(t[1]))
fig, ax = plt.subplots(figsize=(7.2, 4.2))
ax.barh([labels[k] for k, _ in cv], [v for _, v in cv], color=[G if v > 0 else RED for _, v in cv])
ax.axvline(0, color='k', lw=1); ax.set_xlabel('Correlación de rangos de Spearman con el VAN')
ax.set_title('Variables que más explican la variación del VAN', fontweight='bold'); ax.spines[['top', 'right']].set_visible(False)
ax.tick_params(axis='y', labelsize=8)
fig.tight_layout(); fig.savefig(OUT + 'fig_mc_correlacion.png', dpi=200); plt.close(fig)

# mapa de priorización de hipótesis (importancia vs evidencia)
H = [('H1', 'Valora que el pago\nbeneficie al productor', 6, 6.5), ('H2', 'Paga S/ 100+ por\nexperiencia', 5, 5),
     ('H3', 'Productor acepta\nvisitas', 4, 6.5), ('H4', 'Rendimiento\n22 qq/ha', 8, 3), ('H5', 'Preventa\nS/ 120', 10, 1),
     ('H6', 'Productor firma\ncompromiso', 9, 3), ('H7', 'Valora trazabilidad\ncon videos', 6.5, 1), ('H8', 'Recompra\ny visita', 8.5, 1.5)]
json.dump(H, open('hipotesis_puntaje.json', 'w'))
fig, ax = plt.subplots(figsize=(7.4, 5.4))
ax.axhline(5, color='grey', lw=.8); ax.axvline(5, color='grey', lw=.8)
ax.add_patch(plt.Rectangle((0, 5), 5, 5.6, color=RED, alpha=.10)); ax.set_xlim(0, 11); ax.set_ylim(0, 10.6)
ax.text(0.15, 5.15, 'Importante y con poca evidencia: PROBAR PRIMERO', fontsize=8, color=RED, va='bottom')
for k, t, imp, ev in H:
    ax.scatter(ev, imp, s=330, color=(G if ev > 5 else (RED if imp > 5 else A)), zorder=3)
    ax.text(ev, imp, k, ha='center', va='center', color='white', fontweight='bold', zorder=4)
    ax.text(ev + .45, imp, t, fontsize=7, va='center')
ax.set_xlabel('Evidencia disponible (0 = ninguna, 10 = alta)'); ax.set_ylabel('Importancia para el VAN y la decisión (0–10)')
ax.set_title('Mapa de priorización de hipótesis', fontweight='bold'); ax.spines[['top', 'right']].set_visible(False)
fig.tight_layout(); fig.savefig(OUT + 'fig_mapa_hipotesis.png', dpi=200); plt.close(fig)

# organigrama
fig, ax = plt.subplots(figsize=(7.4, 4.4)); ax.axis('off'); ax.set_xlim(0, 11); ax.set_ylim(0, 6)
def box(x, y, t, c=G, w=2.2, h=.8):
    ax.add_patch(FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle='round,pad=0.05', fc=c, ec='none'))
    ax.text(x, y, t, ha='center', va='center', color='white', fontsize=7.5, fontweight='bold')
box(5.5, 5.3, 'Junta de socios\n(4 fundadores)', V, 2.6)
box(5.5, 4.0, 'Gerente general', G, 2.4)
for x, t in [(1.3, 'Jefe agronómico'), (3.7, 'Gestión comercial\ny digital'), (6.1, 'Finanzas y\nadministración'), (8.5, 'Operaciones\ny logística')]:
    ax.plot([5.5, 5.5], [3.6, 3.2], color='k', lw=1); ax.plot([1.3, 8.5], [3.2, 3.2], color='k', lw=1); ax.plot([x, x], [3.2, 2.9], color='k', lw=1)
    box(x, 2.5, t, '#5B7F95', 2.1, .85)
for x, t in [(1.3, 'Técnicos de campo\n(1 por 30 productores)'), (3.7, 'Ejecutivo de ventas\ny atención'), (6.1, 'Contabilidad\nexterna'), (8.5, 'Asistente de almacén y\ncomunicador audiovisual')]:
    ax.plot([x, x], [2.05, 1.7], color='k', lw=1); box(x, 1.2, t, '#8aa07a', 2.3, .85)
ax.text(5.5, 0.35, 'Alianzas y turismo: gerente general con hotel y operadores privados', ha='center', fontsize=7.5, color='#444')
fig.tight_layout(); fig.savefig(OUT + 'fig_organigrama.png', dpi=200); plt.close(fig)
print('ok', len(os.listdir(OUT)))
