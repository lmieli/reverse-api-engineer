"""Exporta produtos e preços da Bramil em Casa (Nova Friburgo) para .xlsx."""
import sys, time
import openpyxl
from openpyxl.styles import Font
from api_client import BramilClient, DEFAULT_CD

IMG = "https://produto-assets-vipcommerce-com-br.br-se1.magaluobjects.com/250x250/"
out = sys.argv[1] if len(sys.argv) > 1 else "bramilemcasa_produtos.xlsx"
c = BramilClient(DEFAULT_CD)
rows, seen = [], set()
for dep in c.departamentos():
    n = 0
    for p in c.produtos_departamento(dep["classificacao_mercadologica_id"]):
        if p["produto_id"] in seen:
            continue
        seen.add(p["produto_id"])
        n += 1
        rows.append([
            dep["descricao"], p["descricao"], float(p["preco"]),
            "Sim" if p["em_oferta"] else "Não",
            float(p["preco_original"]) if p.get("preco_original") else None,
            p.get("unidade_sigla"), p.get("codigo_barras"), p.get("sku"),
            "Sim" if p.get("disponivel") else "Não",
            IMG + p["imagem"] if p.get("imagem") else None,
        ])
    print(dep["descricao"], n, flush=True)

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Produtos"
ws.append(["Departamento", "Produto", "Preço (R$)", "Em oferta", "Preço original (R$)",
           "Unidade", "Código de barras", "SKU", "Disponível", "Imagem (URL)"])
for c_ in ws[1]: c_.font = Font(bold=True)
for r in rows: ws.append(r)
for col in (3, 5):
    for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
        for x in cell: x.number_format = '"R$" #,##0.00'
for k, w in zip("ABCDEFGHIJ", (26, 60, 12, 10, 18, 9, 16, 14, 11, 90)): ws.column_dimensions[k].width = w
ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
wb.save(out); print("total", len(rows))
