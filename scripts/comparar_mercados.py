"""Comparação de um produto entre todos os mercados (Bramil, Superpão, Remercado, Tio Dongo).

Junta os de-paras 1:1 (Bramil x Superpão, Bramil x Remercado, Remercado x Superpão e os do Tio Dongo) em grupos de
produto. Um grupo = um produto em no máximo um item por mercado. Quando os de-paras se contradizem
(dois itens do mesmo mercado no mesmo grupo), mantém o maior conjunto de pares sem contradição.

Uso: python3 scripts/comparar_mercados.py [saida.xlsx]
Saídas: xlsx (Comparação, Resumo), scripts/produto_grupo.csv (grupo -> item de cada mercado).
"""
import csv, collections, itertools, json, os, re, sys
import openpyxl
from openpyxl.styles import Font
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import compare_prices as C

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
S = os.path.join(RAIZ, "scripts")
NOMES = {"B": "Bramil", "S": "Superpão", "R": "Remercado", "T": "Tio Dongo"}


def ler_arestas():
    e = []
    for r in csv.DictReader(open(f"{S}/depara_bramil_superpao.csv", encoding="utf-8")):
        if r["superpao_id"]: e.append((("B", r["bramil_sku"]), ("S", r["superpao_id"])))
    for r in csv.DictReader(open(f"{S}/depara_bramil_remercado.csv", encoding="utf-8")):
        e.append((("B", r["bramil_sku"]), ("R", r["remercado_id"])))
    for r in csv.DictReader(open(f"{S}/depara_remercado_superpao.csv", encoding="utf-8")):
        e.append((("R", r["remercado_id"]), ("S", r["superpao_id"])))
    for arq, a, b, ca, cb in (("depara_bramil_tiodongo.csv", "B", "T", "bramil_sku", "tiodongo_id"),
                              ("depara_tiodongo_superpao.csv", "T", "S", "tiodongo_id", "superpao_id"),
                              ("depara_tiodongo_remercado.csv", "T", "R", "tiodongo_id", "remercado_id")):
        p = f"{S}/{arq}"
        if os.path.exists(p):
            for r in csv.DictReader(open(p, encoding="utf-8")):
                e.append(((a, r[ca]), (b, r[cb])))
    return e


def componentes(arestas):
    par = {}
    def find(x):
        while par.setdefault(x, x) != x:
            par[x] = par[par[x]]; x = par[x]
        return x
    for a, b in arestas: par[find(a)] = find(b)
    comp = collections.defaultdict(set)
    for x in list(par): comp[find(x)].add(x)
    return list(comp.values())


def valido(nos, arestas):
    for c in componentes(arestas):
        if any(v > 1 for v in collections.Counter(m for m, _ in c).values()): return False
    return True


def grupos():
    arestas = ler_arestas()
    out = []
    for c in componentes(arestas):
        if all(v == 1 for v in collections.Counter(m for m, _ in c).values()):
            out.append(c); continue
        ar = [e for e in arestas if e[0] in c]
        melhor = []
        for k in range(len(ar), 0, -1):  # maior subconjunto sem contradição
            for sub in itertools.combinations(ar, k):
                if valido(c, sub): melhor = list(sub); break
            if melhor: break
        out.extend(componentes(melhor))
    return out


def carregar():
    ws = openpyxl.load_workbook(f"{RAIZ}/bramilemcasa_produtos.xlsx").active
    cab = [c.value for c in ws[1]]
    b = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(cab, r)); d.update(C.prep(d["Produto"])); b[d["SKU"]] = d
    def mf(caminho):
        out = {}
        for p in json.load(open(caminho)):
            if not p.get("preco"): continue
            p.update(C.prep_superpao(p)); p["preco_efetivo"] = p["preco_promocional"] or p["preco"]
            out[str(p["id"])] = p
        return out
    return {"B": b, "S": mf(f"{S}/superpao_scraper/superpao_produtos.json"), "R": mf(f"{S}/remercado_scraper/remercado_produtos.json"), "T": tiodongo()}


def tiodongo():
    p = f"{S}/tiodongo/tiodongo_produtos.json"
    out = {}
    if not os.path.exists(p): return out
    for x in json.load(open(p)):
        of = x.get("oferta") or {}
        promo = float(of["preco_oferta"]) if of.get("nome") == "Encarte" and of.get("preco_oferta") else 0
        kg = (x.get("unidade") or "").upper() == "KG" and x.get("preco_original")
        # item por kg: `preco` é o preço da porção estimada; o preço por kg está em `preco_original`
        base_preco = x["preco_original"] if kg else x["preco"]
        if (x.get("unidade") or "").upper() == "KG" and not kg:  # sem preço por kg: usa o peso da porção do nome
            m = re.search(r"aprox\.?\s*([\d.,]+)\s*(kg|g)\b", C.norm(x["nome"]))
            if m:
                g = float(m.group(1).replace(",", ".")) * (1000 if m.group(2) == "kg" else 1)
                if g > 0: base_preco = x["preco"] / (g / 1000)
        d = dict(nome=x["nome"], preco_promocional=promo, preco_efetivo=promo or base_preco, unidade=x.get("unidade"),
                 departamento=x.get("departamento"), url=x.get("url"),
                 imagem=("https://produto-assets-vipcommerce-com-br.br-se1.magaluobjects.com/250x250/" + x["imagem_arquivo"]) if x.get("imagem_arquivo") else None)
        d.update(C.prep_superpao(d))
        if (x.get("unidade") or "").upper() == "KG": d["tam"] = None  # preço já é por kg; o peso no nome é só a porção
        out[str(x["id"])] = d
    return out


def info(m, d):
    if m == "B":
        return dict(nome=d["Produto"], preco=float(d["Preço (R$)"]), unid=(d.get("Unidade") or "").upper(), tam=d["tam"],
                    img=d.get("Imagem (URL)"), dep=d["Departamento"], url="", oferta=d.get("Em oferta") == "Sim")
    return dict(nome=d["nome"], preco=float(d["preco_efetivo"]), unid=(d.get("unidade") or "").upper(), tam=d["tam"],
                img=d.get("imagem"), dep=d.get("departamento"), url=d.get("url", ""), oferta=bool(d.get("preco_promocional")))


def tipo(it):
    kg = it["tam"] is None and (it["unid"] == "KG" or "kg" in C.norm(it["nome"]).split())
    return "kg" if kg else ("pack" if it["tam"] else "?")


def main(saida):
    dados = carregar(); gs = grupos(); linhas = []
    for gi, g in enumerate(sorted(gs, key=lambda c: sorted(c)), 1):
        its = {m: info(m, dados[m][cod]) for m, cod in g if cod in dados[m]}
        if len(its) < 2: continue
        tipos = {m: tipo(i) for m, i in its.items()}
        amb = {m for m, i in its.items() if tipos[m] == "pack" and i["unid"] == "KG"}  # unidade KG com peso no nome: ambíguo
        conv = "kg" in tipos.values() and "pack" in tipos.values() or bool(amb and len(its) > len(amb))
        preco = {}
        if conv:
            ok = {m: (i["preco"] if tipos[m] == "kg" else i["preco"] / (i["tam"] / 1000))
                  for m, i in its.items() if m not in amb and tipos[m] != "?"}
            ref = sorted(ok.values())[len(ok) // 2] if ok else None
            for m, i in its.items():
                if m in amb and ref:
                    c = [i["preco"], i["preco"] / (i["tam"] / 1000)]
                    preco[m] = round(min(c, key=lambda x: abs(__import__("math").log(x / ref))), 2)
                elif m in ok: preco[m] = round(ok[m], 2)
                # item sem peso e sem unidade KG (ex.: bandeja): preço não comparável por kg, fica de fora
        else:
            preco = {m: round(i["preco"], 2) for m, i in its.items()}
        base = "R$/kg (um lado vende por kg)" if conv else "preço da embalagem"
        if len(preco) < 2: continue
        barato = min(preco, key=preco.get); caro = max(preco, key=preco.get)
        dif = round((preco[caro] - preco[barato]) / preco[barato] * 100, 1) if preco[barato] else None
        ref = its.get("B") or its.get("S") or its["R"]
        linhas.append(dict(gid=gi, g=g, its=its, preco=preco, base=base, barato=barato, caro=caro, dif=dif, nome=ref["nome"], n=len(preco)))
    with open(f"{S}/produto_grupo.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["grupo", "mercado", "codigo"])
        for l in linhas:
            for m, cod in sorted(l["g"]): w.writerow([l["gid"], NOMES[m], cod])
    wb = openpyxl.Workbook(); w = wb.active; w.title = "Comparação"
    cab = ["Produto", "Mercados", "Preço Bramil (R$)", "Preço Superpão (R$)", "Preço Remercado (R$)", "Preço Tio Dongo (R$)", "Mais barato", "Mais caro",
           "Diferença mais caro x mais barato (%)", "Base do preço", "Conferir (>100%)", "Nome Bramil", "Nome Superpão", "Nome Remercado", "Nome Tio Dongo",
           "Imagem Bramil", "Imagem Superpão", "Imagem Remercado", "Imagem Tio Dongo", "Link Superpão", "Link Remercado", "Link Tio Dongo"]
    w.append(cab)
    for l in sorted(linhas, key=lambda x: (-x["n"], x["nome"])):
        it = l["its"]; g = lambda m, k: it[m][k] if m in it else None
        w.append([l["nome"], l["n"], l["preco"].get("B"), l["preco"].get("S"), l["preco"].get("R"), l["preco"].get("T"), NOMES[l["barato"]], NOMES[l["caro"]],
                  l["dif"], l["base"], "conferir" if l["dif"] and l["dif"] > 100 else "", g("B", "nome"), g("S", "nome"), g("R", "nome"), g("T", "nome"), g("B", "img"), g("S", "img"), g("R", "img"), g("T", "img"),
                  g("S", "url"), g("R", "url"), g("T", "url")])
    for c in w[1]: c.font = Font(bold=True)
    w.freeze_panes = "B2"; w.auto_filter.ref = w.dimensions
    r = wb.create_sheet("Resumo"); r.append(["Indicador", "Valor"])
    r.append(["Produtos comparáveis (2+ mercados)", len(linhas)])
    for k in (4, 3, 2): r.append([f"Em {k} mercados", sum(1 for l in linhas if l["n"] == k)])
    for m in "BSRT":
        r.append([f"Vezes que {NOMES[m]} é o mais barato", sum(1 for l in linhas if l["barato"] == m)])
    for m in "BSRT":
        com = [l for l in linhas if m in l["preco"] and l["n"] >= 2]
        dif = [(l["preco"][m] - min(l["preco"].values())) / min(l["preco"].values()) * 100 for l in com if min(l["preco"].values())]
        r.append([f"{NOMES[m]}: diferença média para o mais barato (%)", round(sum(dif) / len(dif), 1) if dif else None])
    for c in r[1]: c.font = Font(bold=True)
    r.column_dimensions["A"].width = 56
    wb.save(saida); print(len(linhas), "produtos comparáveis ->", saida)
    return linhas


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(RAIZ, "comparacao_precos_todos_mercados.xlsx"))
