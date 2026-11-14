"""Base histórica de preços (SQLite). Um registro de preço por produto por data de coleta.

Uso:
  python3 scripts/base_historica/base.py carregar [AAAA-MM-DD]   # lê os arquivos atuais e grava o snapshot
  python3 scripts/base_historica/base.py resumo
Mercado novo: escrever um coletor que devolva dicts no formato de `Registro` e chamar `gravar`.
Reexecutar no mesmo dia substitui o snapshot daquele dia (idempotente).
"""
import sys, json, sqlite3, csv, datetime, os
import openpyxl

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..")
DB = os.path.join(RAIZ, "scripts", "base_historica", "precos_historico.db")

ESQUEMA = """
CREATE TABLE IF NOT EXISTS mercado (id TEXT PRIMARY KEY, nome TEXT, site TEXT);
CREATE TABLE IF NOT EXISTS produto (
  mercado TEXT, codigo TEXT, nome TEXT, departamento TEXT, unidade TEXT,
  barras TEXT, imagem TEXT, url TEXT, primeira_coleta TEXT, ultima_coleta TEXT,
  PRIMARY KEY (mercado, codigo));
CREATE TABLE IF NOT EXISTS preco (
  mercado TEXT, codigo TEXT, data TEXT, preco REAL, preco_promocional REAL, preco_original REAL,
  em_oferta INTEGER, disponivel INTEGER,
  PRIMARY KEY (mercado, codigo, data));
CREATE TABLE IF NOT EXISTS coleta (mercado TEXT, data TEXT, itens INTEGER, PRIMARY KEY (mercado, data));
CREATE TABLE IF NOT EXISTS depara (bramil_sku TEXT, superpao_id TEXT, origem TEXT);
CREATE TABLE IF NOT EXISTS depara_remercado (bramil_sku TEXT, remercado_id TEXT, origem TEXT);
CREATE TABLE IF NOT EXISTS depara_remercado_superpao (remercado_id TEXT, superpao_id TEXT, origem TEXT);
CREATE TABLE IF NOT EXISTS produto_grupo (grupo INTEGER, mercado TEXT, codigo TEXT);
CREATE INDEX IF NOT EXISTS ix_preco_data ON preco(data);
"""
MERCADOS = [("bramil", "Bramil em Casa (Nova Friburgo)", "https://www.bramilemcasa.com.br/"),
            ("superpao", "Superpão", "https://www.superpao.com/loja"),
            ("remercado", "Remercado", "https://www.remercado.com.br/loja1"),
            ("tiodongo", "Tio Dongo", "https://www.tiodongo.com.br/")]  # Tio Dongo: 403 fora do Brasil, coletor pendente

def conectar():
    c = sqlite3.connect(DB); c.executescript(ESQUEMA)
    c.executemany("INSERT OR IGNORE INTO mercado VALUES (?,?,?)", MERCADOS)
    return c

def gravar(c, mercado, data, registros):
    """registros: dicts com codigo, nome, departamento, unidade, barras, imagem, url,
    preco, preco_promocional (0/None = sem promo), preco_original, em_oferta, disponivel."""
    c.execute("DELETE FROM preco WHERE mercado=? AND data=?", (mercado, data))
    for r in registros:
        c.execute("""INSERT INTO produto VALUES (?,?,?,?,?,?,?,?,?,?)
          ON CONFLICT(mercado,codigo) DO UPDATE SET nome=excluded.nome, departamento=excluded.departamento,
          unidade=excluded.unidade, barras=COALESCE(excluded.barras,barras), imagem=excluded.imagem,
          url=COALESCE(excluded.url,url), primeira_coleta=MIN(primeira_coleta,excluded.primeira_coleta),
          ultima_coleta=MAX(ultima_coleta,excluded.ultima_coleta)""",
          (mercado, str(r["codigo"]), r["nome"], r.get("departamento"), r.get("unidade"), r.get("barras"),
           r.get("imagem"), r.get("url"), data, data))
        c.execute("INSERT INTO preco VALUES (?,?,?,?,?,?,?,?)",
                  (mercado, str(r["codigo"]), data, r["preco"], r.get("preco_promocional") or None,
                   r.get("preco_original") or None, int(bool(r.get("em_oferta"))), int(r.get("disponivel", 1))))
    c.execute("INSERT OR REPLACE INTO coleta VALUES (?,?,?)", (mercado, data, len(registros)))
    c.commit()

def le_bramil(caminho):
    ws = openpyxl.load_workbook(caminho).active
    cab = [x.value for x in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(cab, row))
        yield dict(codigo=d["SKU"], nome=d["Produto"], departamento=d["Departamento"], unidade=d["Unidade"],
                   barras=d.get("Código de barras"), imagem=d.get("Imagem (URL)"), url=None,
                   preco=d["Preço (R$)"], preco_original=d.get("Preço original (R$)"),
                   em_oferta=d.get("Em oferta") == "Sim", disponivel=1 if d.get("Disponível") in ("Sim", True, 1, None) else 0)

def le_mercafacil(caminho):
    for p in json.load(open(caminho)):
        if not p.get("preco"): continue
        promo = p.get("preco_promocional") or 0
        yield dict(codigo=p["id"], nome=p["nome"], departamento=p.get("departamento"), unidade=p.get("unidade"),
                   imagem=p.get("imagem"), url=p.get("url"), preco=p["preco"], preco_promocional=promo,
                   preco_original=p["preco"] if promo else None, em_oferta=bool(promo),
                   disponivel=1 if (p.get("estoque") or 0) > 0 else 0)

def carregar(data):
    c = conectar()
    gravar(c, "bramil", data, list(le_bramil(os.path.join(RAIZ, "bramilemcasa_produtos.xlsx"))))
    gravar(c, "superpao", data, list(le_mercafacil(os.path.join(RAIZ, "scripts/superpao_scraper/superpao_produtos.json"))))
    rem = os.path.join(RAIZ, "scripts/remercado_scraper/remercado_produtos.json")
    if os.path.exists(rem):
        gravar(c, "remercado", data, list(le_mercafacil(rem)))
    c.execute("DELETE FROM depara")
    for r in csv.DictReader(open(os.path.join(RAIZ, "scripts/depara_bramil_superpao.csv"), encoding="utf-8")):
        c.execute("INSERT INTO depara VALUES (?,?,?)", (r["bramil_sku"], r["superpao_id"] or None, r["origem"]))
    c.execute("DELETE FROM depara_remercado")
    arq = os.path.join(RAIZ, "scripts/depara_bramil_remercado.csv")
    if os.path.exists(arq):
        for r in csv.DictReader(open(arq, encoding="utf-8")):
            c.execute("INSERT INTO depara_remercado VALUES (?,?,?)", (r["bramil_sku"], r["remercado_id"], r["origem"]))
    c.execute("DELETE FROM depara_remercado_superpao")
    arq2 = os.path.join(RAIZ, "scripts/depara_remercado_superpao.csv")
    if os.path.exists(arq2):
        for r in csv.DictReader(open(arq2, encoding="utf-8")):
            c.execute("INSERT INTO depara_remercado_superpao VALUES (?,?,?)", (r["remercado_id"], r["superpao_id"], r["origem"]))
    c.execute("DELETE FROM produto_grupo")
    arq3 = os.path.join(RAIZ, "scripts/produto_grupo.csv")
    if os.path.exists(arq3):
        ids = {"Bramil": "bramil", "Superpão": "superpao", "Remercado": "remercado"}
        for r in csv.DictReader(open(arq3, encoding="utf-8")):
            c.execute("INSERT INTO produto_grupo VALUES (?,?,?)", (int(r["grupo"]), ids[r["mercado"]], r["codigo"]))
    c.commit(); resumo(c)

def resumo(c=None):
    c = c or conectar()
    for m, d, n in c.execute("SELECT mercado,data,itens FROM coleta ORDER BY data,mercado"):
        print(f"{d}  {m:9s} {n} itens")
    print("pares Bramil x Superpao:", c.execute("SELECT COUNT(*) FROM depara WHERE superpao_id IS NOT NULL").fetchone()[0])
    print("pares Bramil x Remercado:", c.execute("SELECT COUNT(*) FROM depara_remercado").fetchone()[0])
    print("produtos comparáveis (grupos):", c.execute("SELECT COUNT(DISTINCT grupo) FROM produto_grupo").fetchone()[0])
    print("pares Remercado x Superpao:", c.execute("SELECT COUNT(*) FROM depara_remercado_superpao").fetchone()[0])

if __name__ == "__main__":
    a = sys.argv[1:] or ["resumo"]
    if a[0] == "carregar": carregar(a[1] if len(a) > 1 else datetime.date.today().isoformat())
    else: resumo()
