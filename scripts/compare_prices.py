"""Compara preços Bramil em Casa (Nova Friburgo) x Superpão.

Uso: python compare_prices.py bramil.xlsx superpao.json saida.xlsx

Pareamento por nome normalizado + tamanho (g/ml/un), sem código de barras em
comum. Cada par recebe uma nota de similaridade (0-100); filtre por "Confiança"
na planilha e confira os pares de nota baixa.
"""
import json, re, sys, unicodedata
from difflib import SequenceMatcher
import openpyxl
from openpyxl.styles import Font, PatternFill

STOP = {"de", "da", "do", "com", "sem", "e", "em", "para", "c", "un", "unid", "unidade",
        "pct", "pacote", "cx", "caixa", "pote", "lata", "lt", "garrafa", "bandeja", "kg", "g",
        "ml", "l", "cong", "congelado", "resf", "tradicional"}


def norm(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9,.\s/]", " ", s)


def tamanho(s):
    """Retorna o tamanho em g ou ml (ou None)."""
    s = norm(s)
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(kg|gr|g|ml|lt|l)\b", s)
    if not m:
        return None
    v = float(m.group(1).replace(",", "."))
    u = m.group(2)
    if u in ("kg", "l", "lt"):
        v *= 1000
    return round(v)


def pack(s):
    """Multiplicador de embalagem (6x350ml, pack, c/12) ou 1."""
    s = norm(s)
    m = re.search(r"\b(\d+)\s*x\s*\d", s) or re.search(r"\b(?:c|com)/?\s*(\d+)\s*(?:un|und|unid)?\b", s)
    if m:
        return int(m.group(1))
    return 6 if re.search(r"\bpack\b", s) else 1


def tokens(s):
    s = re.sub(r"\d+(?:[.,]\d+)?\s*(kg|gr|g|ml|lt|l)\b", " ", norm(s))
    return {t for t in re.findall(r"[a-z0-9]+", s) if t not in STOP and len(t) > 1}


def nota(a, b):
    ta, tb = a["tok"], b["tok"]
    if not ta or not tb:
        return 0
    jac = len(ta & tb) / len(ta | tb)
    cob = len(ta & tb) / min(len(ta), len(tb))
    seq = SequenceMatcher(None, a["n"], b["n"]).ratio()
    return round(100 * (0.4 * jac + 0.4 * cob + 0.2 * seq))


def prep(nome):
    return {"nome": nome, "n": " ".join(sorted(tokens(nome))), "tok": tokens(nome), "tam": tamanho(nome), "pack": pack(nome)}


def main(bramil_xlsx, superpao_json, saida):
    ws = openpyxl.load_workbook(bramil_xlsx).active
    cab = [c.value for c in ws[1]]
    br = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(cab, r))
        d.update(prep(d["Produto"]))
        d["preco"] = d["Preço (R$)"]
        br.append(d)
    sp = [p for p in json.load(open(superpao_json)) if p.get("preco")]
    for p in sp:
        p.update(prep(p["nome"] + " " + (p.get("descricao_quantidade") or "")))
        p["preco_efetivo"] = p["preco_promocional"] if p.get("preco_promocional") else p["preco"]

    idx = {}
    for i, p in enumerate(sp):
        for t in p["tok"]:
            idx.setdefault(t, []).append(i)

    pares, usados = [], set()
    for b in br:
        cand = set()
        for t in b["tok"]:
            cand.update(idx.get(t, []))
        best = (0, None)
        for i in cand:
            p = sp[i]
            if b["pack"] != p["pack"]:
                continue
            if (b["tam"] is None) != (p["tam"] is None) or b["tam"] != p["tam"]:
                continue
            s = nota(b, p)
            if s > best[0]:
                best = (s, i)
        if best[1] is not None and best[0] >= 80:
            pares.append((b, sp[best[1]], best[0]))
            usados.add(best[1])

    wb = openpyxl.Workbook()
    w = wb.active; w.title = "Comparação"
    w.append(["Produto Bramil", "Produto Superpão", "Preço Bramil (R$)", "Preço Superpão (R$)",
              "Diferença (R$) Bramil-Superpão", "Bramil é mais barato (%)", "Mais barato", "Confiança",
              "Depto Bramil", "Depto Superpão", "Link Superpão", "Conferir"])
    for b, p, s in sorted(pares, key=lambda x: -x[2]):
        pb, ps = float(b["preco"]), float(p["preco_efetivo"])
        mais = "Bramil" if pb < ps else "Superpão" if ps < pb else "Empate"
        w.append([b["Produto"], p["nome"], pb, ps, round(pb - ps, 2),
                  round((ps - pb) / ps * 100, 1) if ps else None, mais, s,
                  b["Departamento"], p["departamento"], p["url"],
                  "Conferir (diferença > 40%)" if ps and abs(pb - ps) / ps > 0.4 else ""])
    for c in w[1]: c.font = Font(bold=True)
    for col in "CDE":
        for c in w[col][1:]: c.number_format = '"R$" #,##0.00'
    for k, v in zip("ABCDEFGHIJKL", (48, 48, 14, 14, 18, 16, 12, 11, 24, 18, 60, 26)): w.column_dimensions[k].width = v
    w.freeze_panes = "A2"; w.auto_filter.ref = w.dimensions
    verde = PatternFill("solid", fgColor="E2F0D9")
    for row in w.iter_rows(min_row=2):
        if row[6].value == "Bramil": row[6].fill = verde

    w2 = wb.create_sheet("Superpão")
    w2.append(["Produto", "Marca", "Departamento", "Preço (R$)", "Preço promocional (R$)", "Unidade", "Estoque", "Imagem", "Link"])
    for p in sp:
        w2.append([p["nome"], p["marca"], p["departamento"], p["preco"], p["preco_promocional"], p["unidade"],
                   p["estoque"], p["imagem"], p["url"]])
    w3 = wb.create_sheet("Bramil")
    w3.append(cab)
    for b in br: w3.append([b[c] for c in cab])
    for s_ in (w2, w3):
        for c in s_[1]: c.font = Font(bold=True)
        s_.freeze_panes = "A2"; s_.auto_filter.ref = s_.dimensions
    wb.save(saida)
    print(f"pares: {len(pares)} de {len(br)} Bramil / {len(sp)} Superpão")


if __name__ == "__main__":
    main(*sys.argv[1:4])
