"""Extrai produtos e preços públicos do Remercado (https://www.remercado.com.br/loja1).

Plataforma Mercafacil (Next.js). A API GraphQL exige token, então este script
lê apenas as páginas públicas de categoria listadas no sitemap.xml (permitidas
pelo robots.txt), cujo HTML já traz os produtos renderizados no servidor.
"""
import json, re, sys, time, urllib.request

BASE = "https://www.remercado.com.br"
UA = {"user-agent": "Mozilla/5.0"}
PAUSA = 0.3


def get(url):
    for tentativa in range(3):
        try:
            req = urllib.request.Request(url, headers=UA)
            return urllib.request.urlopen(req, timeout=30).read().decode()
        except Exception:
            time.sleep(2 * (tentativa + 1))
    return ""


def rsc(url):
    html = get(url)
    partes = re.findall(r'self\.__next_f\.push\(\[1,"(.*?)"\]\)</script>', html, re.S)
    return "".join(json.loads('"' + p + '"') for p in partes)


def categorias():
    sm = get(BASE + "/sitemap.xml")
    return [u for u in re.findall(r"<loc>([^<]+)", sm)
            if "/produto/" not in u and re.search(r"-\d+$", u)]


def produtos_da_pagina(texto):
    dec = json.JSONDecoder()
    out = []
    for m in re.finditer(r'\{"__typename":"EcommerceProduct"', texto):
        try:
            obj, _ = dec.raw_decode(texto, m.start())
            out.append(obj)
        except ValueError:
            pass
    return out


def main(saida):
    vistos, linhas = set(), []
    cats = categorias()
    for i, url in enumerate(cats, 1):
        pagina, total = 1, 1
        while pagina <= total:
            t = rsc(url if pagina == 1 else f"{url}?page={pagina}")
            tp = re.search(r'"totalPages":(\d+)', t)
            total = int(tp.group(1)) if tp else 1
            for p in produtos_da_pagina(t):
                if p["id"] in vistos:
                    continue
                vistos.add(p["id"])
                linhas.append({
                    "id": p["id"], "nome": p["name"], "marca": p.get("brand"),
                    "departamento": p.get("department"),
                    "categoria": url.rsplit("/", 1)[-1],
                    "preco": p["price"], "preco_promocional": p.get("priceWithDiscount"),
                    "desconto_pct": p.get("discount"),
                    "unidade": (p.get("unitOfMeasurement") or {}).get("abbreviation"),
                    "descricao_quantidade": p.get("quantityDescription"),
                    "estoque": p.get("stock"), "imagem": p.get("image"),
                    "url": f"{BASE}/loja1/produto/m/{p['slug']}-{p['id']}",
                })
            pagina += 1
            time.sleep(PAUSA)
        print(f"{i}/{len(cats)} {url.rsplit('/',1)[-1]} -> {len(linhas)}", flush=True)
    json.dump(linhas, open(saida, "w"), ensure_ascii=False, indent=1)
    print("total", len(linhas))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "remercado_produtos.json")
