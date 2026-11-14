"""Compara preços entre Bramil em Casa e Superpão com pareamento 1 para 1.

Uso: python compare_prices.py bramil.xlsx superpao.json saida.xlsx

Regras do pareamento (de-para):
  * 1 para 1: cada produto de um mercado tem no máximo um par no outro, e
    nenhum produto é usado duas vezes. Quem não tem equivalente fica sem par.
  * Tamanho (g/ml) e embalagem (pack 6x, c/12) precisam ser iguais.
  * Variações diferentes (ex.: vinagre de maçã x de álcool) não pareiam: se os
    dois lados têm palavras que o outro não tem, o par só é aceito com nota alta.
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
        "gr", "grt", "gratis"}
NOTA_MIN = 75          # nota mínima de qualquer par
NOTA_CONFLITO = 85     # nota mínima quando os dois lados têm palavras sem correspondente


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
    s = norm(s)
    m = re.search(r"\b(\d+)\s*x\s*\d", s) or re.search(r"\b(?:c|com)/?\s*(\d+)\s*(?:un|und|unid)?\b", s)
    if m:
        return int(m.group(1))
    return 6 if re.search(r"\bpack\b", s) else 1


def tokens(s):
    s = re.sub(r"\d+(?:[.,]\d+)?\s*(kg|gr|g|ml|lt|l)\b", " ", norm(s))
    s = re.sub(r"\b\d+\s*x\b", " ", s)
    out = []
    for t in re.findall(r"[a-z0-9]+", s):
        if t not in STOP and len(t) > 1 and t not in out:
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
    """Retorna (nota 0-100, sobras_a, sobras_b)."""
    ta, tb = sorted(a["tok"], key=lambda t: -peso[t]), list(b["tok"])
    usados, casado = set(), 0.0
    sobra_a = []
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
            casado += melhor * (peso[t] + peso[tb[mi]]) / 2
    sobra_b = [u for j, u in enumerate(tb) if j not in usados]
    wa = sum(peso[t] for t in a["tok"])
    wb = sum(peso[t] for t in b["tok"])
    if wa + wb == 0:
        return 0, sobra_a, sobra_b
    return round(100 * 2 * casado / (wa + wb)), sobra_a, sobra_b


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
            if b["pack"] != p["pack"] or b["tam"] != p["tam"]:
                continue
            nota, sa, sb = pontuar(b, p, peso)
            minimo = NOTA_CONFLITO if (sa and sb) else NOTA_MIN
            if nota >= minimo:
                candidatos.append((nota, -len(sa) - len(sb), bi, i))

    candidatos.sort(reverse=True)
    pb, ps, pares = set(), set(), []
    for nota, _, bi, i in candidatos:
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
