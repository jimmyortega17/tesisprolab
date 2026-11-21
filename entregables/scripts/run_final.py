import sys, re; sys.path.insert(0, '.')
import build_final as B
from pages import to_pdf, compute
from docx import Document

def collect(path):
    d = Document(path); ps = d.paragraphs
    j = next(i for i, p in enumerate(ps) if p.style.name == 'Heading 1' and p.text.startswith('Capítulo I:'))
    entries = []; ap = False; capt = []; capf = []
    L = ps[j:]
    for k, p in enumerate(L):
        if p.style.name == 'Heading 1': entries.append(('ch', p.text.strip())); ap = p.text.strip() == 'Apéndices'
        elif p.style.name == 'Heading 2': entries.append(('ap' if ap else 'sub', p.text.strip()))
        m = re.match(r'^(Tabla|Figura) (\d+)$', p.text.strip())
        if m and p.runs and p.runs[0].bold: (capt if m.group(1) == 'Tabla' else capf).append((int(m.group(2)), L[k + 1].text.strip()))
    return entries, capt, capf

B.build(None)
e, t, f = collect(B.OUT)
pdf = to_pdf(B.OUT); pages, start, n = compute(pdf, e, t, f)
B.build(pages)
pdf = to_pdf(B.OUT); pages2, start2, n2 = compute(pdf, e, t, f)
dif = {k: (pages[k], pages2[k]) for k in pages if pages2.get(k) != pages[k]}
print('páginas pdf', n2, 'diferencias entre pasadas', len(dif), list(dif.items())[:5])
if dif:
    B.build(pages2); pdf = to_pdf(B.OUT); pages3, _, _ = compute(pdf, e, t, f)
    print('tercera pasada diferencias', sum(1 for k in pages2 if pages3.get(k) != pages2[k]))
print(len(e), 'entradas;', len(t), 'tablas;', len(f), 'figuras')
