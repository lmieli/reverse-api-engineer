"""Compara preços entre Bramil em Casa e Superpão com pareamento 1 para 1.

Uso: python compare_prices.py bramil.xlsx superpao.json saida.xlsx

Regras do pareamento (de-para):
  * 1 para 1: cada produto de um mercado tem no máximo um par no outro, e
    nenhum produto é usado duas vezes. Quem não tem equivalente fica sem par.
  * Tamanho (g/ml) e embalagem (pack 6x, c/12) precisam ser iguais.
  * Variações diferentes (ex.: vinagre de maçã x de álcool) não pareiam: as palavras
    mais distintivas (marca, sabor, tipo) precisam ter correspondente; só palavras
    genéricas ("bala" x "pastilha", "licor" x "whisky") podem divergir.
  * A atribuição é global: os pares mais certos são fechados primeiro, e quem
    perdeu o seu candidato fica sem par em vez de herdar um parecido.
Sem código de barras em comum, a nota (0-100) indica a confiança do par.
"""
import json, math, re, sys, unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher
from functools import lru_cache
import openpyxl
from openpyxl.styles import Font, PatternFill

STOP = {"de", "da", "do", "das", "dos", "com", "sem", "e", "em", "para", "c", "un", "unid", "unidade",
        "und", "pct", "pacote", "cx", "caixa", "pote", "lata", "lt", "garrafa", "bandeja", "kg", "g",
        "ml", "l", "cong", "congelado", "resf", "tradicional", "trad", "tp", "vd", "pet", "lit", "litro",
        "gr", "grt", "gratis", "grama", "gramas", "litro", "litros", "unidades", "uni", "no", "nr",
        "aprox", "aproximadamente"}
NOTA_MIN = 70          # nota mínima de qualquer par
NOTA_SEM_NUCLEO = 85   # com nota alta, o núcleo não é exigido
NUCLEO = 2             # as N palavras mais distintivas do produto menor precisam casar


def norm(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9,.\s/]", " ", s)


def tamanho(s):
    """Tamanho em g ou ml (None se não houver)."""
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(kg|gr|g|ml|lt|l)\b", norm(s))
    if not m:
        return None
    v = float(m.group(1).replace(",", "."))
    if m.group(2) in ("kg", "l", "lt"):
        v *= 1000
    return round(v)


def pack(s):
    """Quantidade de unidades na embalagem (6x350ml, c/12, 2un, leve 3) ou 1."""
    s = norm(s)
    m = (re.search(r"\b(\d+)\s*x\s*\d", s) or re.search(r"\bc/\s*(\d+)", s)
         or re.search(r"\bcom\s+(\d+)\s*(?:un|und|unid)", s) or re.search(r"\b(\d+)\s*(?:un|und|unid|unidades)\b", s))
    if m:
        return max(1, int(m.group(1)))
    return 6 if re.search(r"\bpack\b", s) else 1


def stem(t):
    return t[:-1] if len(t) > 4 and t.endswith("s") and not t.endswith("ss") else t


def tokens(s):
    s = re.sub(r"\d+(?:[.,]\d+)?\s*(kg|gr|g|ml|lt|l|un|und|unid|unidades)\b", " ", norm(s))
    s = re.sub(r"\b\d+\s*x\b|\bc/\s*\d+", " ", s)
    out = []
    for t in re.findall(r"[a-z0-9]+", s):
        t = stem(t)
        if t not in STOP and len(t) > 1 and not t.isdigit() and t not in out:
            out.append(t)
    return out


@lru_cache(maxsize=None)
def sim(a, b):
    if a == b:
        return 1.0
    if min(len(a), len(b)) >= 3 and (a.startswith(b) or b.startswith(a)):
        return 0.9
    if abs(len(a) - len(b)) <= 2 and min(len(a), len(b)) >= 4:
        r = SequenceMatcher(None, a, b).ratio()
        return r if r >= 0.85 else 0.0
    return 0.0


def prep(nome):
    return {"tok": tokens(nome), "tam": tamanho(nome), "pack": pack(nome)}


def pontuar(a, b, peso):
    """Retorna (nota 0-100, tokens casados de a, tokens casados de b)."""
    ta, tb = sorted(a["tok"], key=lambda t: -peso[t]), list(b["tok"])
    usados, casado = set(), 0.0
    sobra_a, casados_a = [], set()
    for t in ta:
        melhor, mi = 0.0, None
        for j, u in enumerate(tb):
            if j in usados:
                continue
            s = sim(t, u)
            if s > melhor:
                melhor, mi = s, j
        if mi is None:
            sobra_a.append(t)
        else:
            usados.add(mi)
            casados_a.add(t)
            casado += melhor * (peso[t] + peso[tb[mi]]) / 2
    sobra_b = [u for j, u in enumerate(tb) if j not in usados]
    casados_b = {u for j, u in enumerate(tb) if j in usados}
    wa = sum(peso[t] for t in a["tok"])
    wb = sum(peso[t] for t in b["tok"])
    if wa + wb == 0:
        return 0, set(), set()
    return round(100 * 2 * casado / (wa + wb)), casados_a, casados_b


def mesmo_tamanho(a, b):
    """Tamanhos iguais (tolerância de 2%); os dois ausentes também contam como iguais."""
    if a is None or b is None:
        return True  # um lado sem tamanho: decidido pela nota (exige 90+)
    return abs(a - b) <= max(1, 0.02 * max(a, b))


def parear(br, sp):
    """Atribuição 1 para 1, fechando primeiro os pares de maior nota."""
    df = Counter()
    for x in br + sp:
        df.update(set(x["tok"]))
    n = len(br) + len(sp)
    peso = defaultdict(lambda: 1.0, {t: math.log(n / c) + 1 for t, c in df.items()})

    idx = defaultdict(list)
    for i, p in enumerate(sp):
        for t in p["tok"]:
            idx[t[:4]].append(i)

    candidatos = []
    for bi, b in enumerate(br):
        cont = Counter()
        for t in b["tok"]:
            for i in set(idx.get(t[:4], ())):
                cont[i] += peso[t]
        for i, _ in cont.most_common(60):
            p = sp[i]
            if b["pack"] != p["pack"] or not mesmo_tamanho(b["tam"], p["tam"]):
                continue
            nota, ca, cb = pontuar(b, p, peso)
            if nota < NOTA_MIN or (b["tam"] is None) != (p["tam"] is None) and nota < 90:
                continue
            # núcleo: as palavras mais distintivas do produto com menos palavras (marca,
            # sabor, tipo) precisam ter correspondente; só palavras genéricas podem sobrar
            menor, casados = (b, ca) if len(b["tok"]) <= len(p["tok"]) else (p, cb)
            nucleo = sorted(menor["tok"], key=lambda t: -peso[t])[:NUCLEO]
            if nota < NOTA_SEM_NUCLEO and any(t not in casados for t in nucleo):
                continue
            candidatos.append((nota, bi, i))

    candidatos.sort(reverse=True)
    pb, ps, pares = set(), set(), []
    for nota, bi, i in candidatos:
        if bi in pb or i in ps:
            continue
        pb.add(bi); ps.add(i)
        pares.append((bi, i, nota))
    return pares


def main(bramil_xlsx, superpao_json, saida):
    ws = openpyxl.load_workbook(bramil_xlsx).active
    cab = [c.value for c in ws[1]]
    br = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(cab, r))
        d.update(prep(d["Produto"]))
        br.append(d)
    sp = [p for p in json.load(open(superpao_json)) if p.get("preco")]
    for p in sp:
        p.update(prep(p["nome"] + " " + (p.get("descricao_quantidade") or "")))
        p["preco_efetivo"] = p["preco_promocional"] if p.get("preco_promocional") else p["preco"]

    pares = parear(br, sp)
    par_b = {bi: (i, n) for bi, i, n in pares}
    par_s = {i: (bi, n) for bi, i, n in pares}
    assert len(par_b) == len(par_s) == len(pares)  # garante 1 para 1

    wb = openpyxl.Workbook()
    w = wb.active; w.title = "Comparação"
    w.append(["Produto Bramil", "Produto Superpão", "Preço Bramil (R$)", "Preço Superpão (R$)",
              "Diferença (R$) Bramil-Superpão", "Bramil é mais barato (%)", "Mais barato", "Confiança",
              "Depto Bramil", "Depto Superpão", "Link Superpão", "Conferir",
              "Imagem Bramil (URL)", "Imagem Superpão (URL)"])
    for bi, i, nota in sorted(pares, key=lambda x: -x[2]):
        b, p = br[bi], sp[i]
        pb, ps = float(b["Preço (R$)"]), float(p["preco_efetivo"])
        mais = "Bramil" if pb < ps else "Superpão" if ps < pb else "Empate"
        w.append([b["Produto"], p["nome"], pb, ps, round(pb - ps, 2),
                  round((ps - pb) / ps * 100, 1) if ps else None, mais, nota,
                  b["Departamento"], p["departamento"], p["url"],
                  "Conferir (diferença > 40%)" if ps and abs(pb - ps) / ps > 0.4 else "",
                  b.get("Imagem (URL)"), p.get("imagem")])
    for c in w[1]: c.font = Font(bold=True)
    for col in "CDE":
        for c in w[col][1:]: c.number_format = '"R$" #,##0.00'
    for k, v in zip("ABCDEFGHIJKLMN", (48, 48, 14, 14, 18, 16, 12, 11, 24, 18, 60, 26, 60, 60)):
        w.column_dimensions[k].width = v
    w.freeze_panes = "A2"; w.auto_filter.ref = w.dimensions
    verde = PatternFill("solid", fgColor="E2F0D9")
    for row in w.iter_rows(min_row=2):
        if row[6].value == "Bramil": row[6].fill = verde

    w2 = wb.create_sheet("Superpão")
    w2.append(["Produto", "Marca", "Departamento", "Preço (R$)", "Preço promocional (R$)", "Unidade", "Estoque",
               "Imagem", "Link", "Par na Bramil"])
    for i, p in enumerate(sp):
        par = br[par_s[i][0]]["Produto"] if i in par_s else "Sem par"
        w2.append([p["nome"], p["marca"], p["departamento"], p["preco"], p["preco_promocional"], p["unidade"],
                   p["estoque"], p["imagem"], p["url"], par])
    w3 = wb.create_sheet("Bramil")
    w3.append(cab + ["Par no Superpão"])
    for bi, b in enumerate(br):
        w3.append([b[c] for c in cab] + [sp[par_b[bi][0]]["nome"] if bi in par_b else "Sem par"])
    for s_ in (w2, w3):
        for c in s_[1]: c.font = Font(bold=True)
        s_.freeze_panes = "A2"; s_.auto_filter.ref = s_.dimensions
    wb.save(saida)
    print(f"pares 1 para 1: {len(pares)} | Bramil sem par: {len(br)-len(pares)} de {len(br)} | "
          f"Superpão sem par: {len(sp)-len(pares)} de {len(sp)}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
