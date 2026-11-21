import subprocess, re, os
SCR = '/tmp/claude-0/-home-user-tesisprolab/45d5a49c-f407-5ad1-bffb-270b39202e9f/scratchpad/pdf/'
os.makedirs(SCR, exist_ok=True)
def to_pdf(docx):
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', SCR, docx], capture_output=True, timeout=300)
    return SCR + os.path.basename(docx).replace('.docx', '.pdf')
def page_texts(pdf):
    n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout).group(1))
    out = []
    for i in range(1, n + 1):
        t = subprocess.run(['pdftotext', '-f', str(i), '-l', str(i), '-layout', pdf, '-'], capture_output=True, text=True).stdout
        out.append(t)
    return out
def norm(s): return re.sub(r'\s+', ' ', s).strip()
def compute(pdf, entries, caps_t, caps_f):
    txt = page_texts(pdf)
    num = []
    for t in txt:
        ls = [norm(l) for l in t.split('\n') if l.strip()]
        num.append(int(ls[0]) if ls and ls[0].isdigit() else None)
    start = next(i for i, n in enumerate(num) if n == 1)
    L = [[norm(l) for l in t.split('\n') if l.strip()] for t in txt]
    def find(pred):
        for i in range(start, len(txt)):
            if pred(L[i]): return num[i]
        return None
    pages = {}
    for kind, t in entries:
        key = norm(t)
        def pred(ls, key=key):
            for k, l in enumerate(ls):
                if l == key: return True
                if len(key) > 40 and l.startswith(key[:40]): return True
                if key.startswith(l) and len(l) > 20 and k + 1 < len(ls) and norm(l + ' ' + ls[k + 1]).startswith(key[:len(l) + 6]) and len(key) > len(l): return True
            return False
        p = find(pred)
        if p is None: print('NO ENCONTRADO', t)
        else: pages[('h', t)] = p
    for n, t in caps_t:
        p = find(lambda ls: f'Tabla {n}' in ls)
        if p is None: print('sin tabla', n)
        else: pages[('Tabla', n)] = p
    for n, t in caps_f:
        p = find(lambda ls: f'Figura {n}' in ls)
        if p is None: print('sin figura', n)
        else: pages[('Figura', n)] = p
    return pages, start, len(txt)
