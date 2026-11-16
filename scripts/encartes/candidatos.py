"""Sugere, para cada oferta de encarte, o produto mais parecido em cada mercado (para revisão)."""
import csv, json, math, os, sys
from collections import Counter, defaultdict
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import compare_prices as C, comparar_mercados as M
dados = M.carregar()
pool = {m: [(cod, M.info(m, d)["nome"], M.info(m, d)["preco"], d) for cod, d in dados[m].items()] for m in "BSRT"}
ofertas = list(csv.DictReader(open(os.path.join(os.path.dirname(__file__), "supermarket.csv"), encoding="utf-8")))
for o in ofertas: o.update(C.prep(o["produto"]))
df = Counter()
for m in pool:
    for _, _, _, d in pool[m]: df.update(set(d["tok"]))
n = sum(len(v) for v in pool.values()); peso = defaultdict(lambda: 1.0, {t: math.log(n / c) + 1 for t, c in df.items()})
for o in ofertas:
    print(f"## {o['id']} | {o['produto']} | {o['preco']}")
    for m in "BSRT":
        best = []
        for cod, nome, preco, d in pool[m]:
            if not set(d["tok"]) & set(o["tok"]): continue
            if o["tam"] and d["tam"] and not C.mesmo_tamanho(o["tam"], d["tam"]): continue
            nota, _, _ = C.pontuar(o, d, peso)
            if nota >= 55: best.append((nota, cod, nome, preco))
        best.sort(reverse=True)
        for nota, cod, nome, preco in best[:2]: print(f"   {m}{'ab'[best.index((nota,cod,nome,preco))]} {cod} | {nome[:58]} | {preco} ({nota})")
