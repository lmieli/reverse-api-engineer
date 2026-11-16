"""Compara ofertas de encarte (WhatsApp) com os preços atuais dos quatro mercados.

Entradas: scripts/encartes/<loja>.csv (ofertas transcritas) e <loja>_depara.csv (oferta -> produto por mercado,
tipo exato|variante; "variante" = oferta vale para vários sabores/fragrâncias e foi ligada a um deles).
Saídas: relatorio/encarte_<loja>.xlsx e tabelas encarte / encarte_depara na base histórica.
Uso: python3 scripts/encartes/comparar_encarte.py [loja]   (padrão: supermarket)
"""
import csv, os, sqlite3, sys
import openpyxl
from openpyxl.styles import Font, PatternFill
AQUI = os.path.dirname(os.path.abspath(__file__)); RAIZ = os.path.dirname(os.path.dirname(AQUI))
sys.path.insert(0, os.path.dirname(AQUI))
import comparar_mercados as M

NOMES = {"B": "Bramil", "S": "Superpão", "R": "Remercado", "T": "Tio Dongo"}
DB = os.path.join(os.path.dirname(AQUI), "base_historica", "precos_historico.db")


def num(x):
    return float(x) if x not in (None, "") else None


def main(loja="supermarket"):
    ofertas = list(csv.DictReader(open(f"{AQUI}/{loja}.csv", encoding="utf-8")))
    dp = list(csv.DictReader(open(f"{AQUI}/{loja}_depara.csv", encoding="utf-8")))
    dados = M.carregar()
    por = {}
    for r in dp: por.setdefault(r["oferta_id"], {})[r["mercado"]] = r
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Encarte x mercados"
    cab = ["Oferta", "Produto (encarte)", "Tamanho", "Preço encarte", "Preço condicional", "Condição", "Preço clube",
           "Validade", "Melhor preço encarte"]
    for m in "BSRT": cab += [f"{NOMES[m]} produto", f"{NOMES[m]} preço", f"{NOMES[m]} tipo"]
    cab += ["Mais barato nos mercados", "Preço mais barato", "Encarte vs mais barato (%)", "Fonte"]
    ws.append(cab)
    for c in ws[1]: c.font = Font(bold=True)
    verde = PatternFill("solid", fgColor="C6EFCE"); verm = PatternFill("solid", fgColor="FFC7CE")
    linhas = []; sem = 0
    for o in ofertas:
        pe = min(p for p in (num(o["preco"]), num(o["preco_condicional"]), num(o["preco_clube"])) if p is not None)
        row = [int(o["id"]), o["produto"], o["tamanho"], num(o["preco"]), num(o["preco_condicional"]), o["condicao"],
               num(o["preco_clube"]), f'{o["validade_inicio"]} a {o["validade_fim"]}', pe]
        precos = {}
        for m in "BSRT":
            r = por.get(o["id"], {}).get(m)
            d = dados[m].get(r["codigo"]) if r else None
            if d:
                i = M.info(m, d); precos[m] = i["preco"]; row += [i["nome"], i["preco"], r["tipo"]]
                linhas.append((loja, o["id"], m, r["codigo"], r["tipo"]))
            else: row += [None, None, None]
        if precos:
            mb = min(precos, key=precos.get)
            row += [NOMES[mb], precos[mb], round((pe / precos[mb] - 1) * 100, 1)]
        else: row += [None, None, None]; sem += 1
        row.append(f'{o["fonte"]} ({o["imagem"]})'); ws.append(row)
        if precos:
            ws.cell(ws.max_row, len(cab) - 1).fill = verde if pe <= precos[mb] else verm
    ws.freeze_panes = "C2"; ws.auto_filter.ref = ws.dimensions
    for col, wd in zip("ABCDEFGHI", (7, 50, 10, 10, 10, 22, 10, 22, 10)): ws.column_dimensions[col].width = wd
    os.makedirs(f"{RAIZ}/relatorio", exist_ok=True)
    saida = f"{RAIZ}/relatorio/encarte_{loja}.xlsx"; wb.save(saida)
    c = sqlite3.connect(DB)
    c.executescript("""CREATE TABLE IF NOT EXISTS encarte (loja TEXT, id TEXT, produto TEXT, tamanho TEXT, preco REAL,
        preco_condicional REAL, condicao TEXT, preco_clube REAL, validade_inicio TEXT, validade_fim TEXT, fonte TEXT,
        imagem TEXT, PRIMARY KEY (loja, id));
        CREATE TABLE IF NOT EXISTS encarte_depara (loja TEXT, oferta_id TEXT, mercado TEXT, codigo TEXT, tipo TEXT);""")
    c.execute("DELETE FROM encarte WHERE loja=?", (loja,)); c.execute("DELETE FROM encarte_depara WHERE loja=?", (loja,))
    c.executemany("INSERT INTO encarte VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                  [(loja, o["id"], o["produto"], o["tamanho"], num(o["preco"]), num(o["preco_condicional"]), o["condicao"],
                    num(o["preco_clube"]), o["validade_inicio"], o["validade_fim"], o["fonte"], o["imagem"]) for o in ofertas])
    c.executemany("INSERT INTO encarte_depara VALUES (?,?,?,?,?)", linhas); c.commit()
    com = len(ofertas) - sem
    print(f"{saida}: {len(ofertas)} ofertas, {com} com par em ao menos um mercado, {sem} sem par, {len(linhas)} ligações")


if __name__ == "__main__":
    main(*sys.argv[1:])
