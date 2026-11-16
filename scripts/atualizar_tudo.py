"""Atualização diária: coleta os 4 mercados, grava na base histórica, refaz a comparação e a página.

Uso: python3 scripts/atualizar_tudo.py [--so mercado ...] [--data AAAA-MM-DD]
Mercados: bramil, superpao, remercado, tiodongo.

- Cada coletor grava num arquivo temporário. Só substitui o arquivo atual se a coleta trouxe pelo menos 80% dos
  produtos da anterior; senão mantém o arquivo antigo e não grava aquele mercado na base do dia.
- Os de-paras (CSV) não mudam aqui: produto novo entra na base, mas só entra na comparação depois de pareado.
- Saídas: comparacao_precos_todos_mercados.xlsx, relatorio/precos_mercados.html e um resumo em
  relatorio/ultima_atualizacao.json.
"""
import argparse, datetime, json, os, shutil, subprocess, sys, time
from concurrent.futures import ThreadPoolExecutor

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
S = os.path.join(RAIZ, "scripts")
TMP = os.path.join(RAIZ, ".coleta_tmp")

COLETORES = {
    "bramil": dict(cmd=[sys.executable, "export_xlsx.py"], cwd=f"{S}/bramilemcasa_api", env={"PYTHONPATH": "."},
                   final=f"{RAIZ}/bramilemcasa_produtos.xlsx", ext="xlsx"),
    "superpao": dict(cmd=[sys.executable, "scraper.py"], cwd=f"{S}/superpao_scraper",
                     final=f"{S}/superpao_scraper/superpao_produtos.json", ext="json"),
    "remercado": dict(cmd=[sys.executable, "scraper.py"], cwd=f"{S}/remercado_scraper",
                      final=f"{S}/remercado_scraper/remercado_produtos.json", ext="json"),
    "tiodongo": dict(cmd=["node", "scraper.mjs"], cwd=f"{S}/tiodongo",
                     final=f"{S}/tiodongo/tiodongo_produtos.json", ext="json"),
}


def contar(caminho, ext):
    if not os.path.exists(caminho): return 0
    if ext == "json": return len(json.load(open(caminho)))
    import openpyxl
    return openpyxl.load_workbook(caminho, read_only=True).active.max_row - 1


def coletar(nome):
    c = COLETORES[nome]; saida = f"{TMP}/{nome}.{c['ext']}"
    env = dict(os.environ, **c.get("env", {}))
    t0 = time.time()
    try:
        r = subprocess.run(c["cmd"] + [saida], cwd=c["cwd"], env=env, capture_output=True, text=True, timeout=2 * 3600)
        ok_proc = r.returncode == 0
        erro = (r.stderr or "")[-500:]
    except subprocess.TimeoutExpired:
        ok_proc, erro = False, "tempo esgotado (2h)"
    novo, antigo = contar(saida, c["ext"]), contar(c["final"], c["ext"])
    ok = ok_proc and novo > 0 and (antigo == 0 or novo >= 0.8 * antigo)
    if ok: shutil.move(saida, c["final"])
    return nome, dict(ok=ok, produtos=novo, anterior=antigo, minutos=round((time.time() - t0) / 60, 1),
                      erro=None if ok else (erro or f"coleta trouxe {novo} produtos, anterior {antigo}"))


def pagina(data):
    import openpyxl
    wb = openpyxl.load_workbook(f"{RAIZ}/comparacao_precos_todos_mercados.xlsx", read_only=True)
    inv = {"Bramil": "B", "Superpão": "S", "Remercado": "R", "Tio Dongo": "T"}
    it = wb["Comparação"].iter_rows(values_only=True); cab = list(next(it)); rows = []
    for r in it:
        d = dict(zip(cab, r))
        rows.append([d["Produto"], d["Mercados"], d["Preço Bramil (R$)"], d["Preço Superpão (R$)"], d["Preço Remercado (R$)"],
                     d["Preço Tio Dongo (R$)"], inv[d["Mais barato"]], d["Diferença mais caro x mais barato (%)"], d["Base do preço"],
                     d["Conferir (>100%)"] or "", d["Nome Bramil"], d["Nome Superpão"], d["Nome Remercado"], d["Nome Tio Dongo"],
                     d["Link Superpão"], d["Link Remercado"], d["Link Tio Dongo"], d["Imagem Bramil"], d["Imagem Superpão"],
                     d["Imagem Remercado"], d["Imagem Tio Dongo"]])
    t = open(f"{S}/pagina/modelo.html", encoding="utf-8").read()
    t = t.replace("__COLETA__", datetime.date.fromisoformat(data).strftime("%d/%m/%Y"))
    t = t.replace("__DATA__", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
    enc = []
    if os.path.exists(f"{RAIZ}/relatorio/encarte_supermarket.xlsx"):
        it = openpyxl.load_workbook(f"{RAIZ}/relatorio/encarte_supermarket.xlsx", read_only=True).active.iter_rows(values_only=True)
        next(it)
        for r in it:  # id, produto, tam, preço, cond, condição, clube, validade, melhor, 4x(nome, preço, tipo), barato, preço, %, fonte
            r = list(r); enc.append(r[:2] + r[3:21] + [inv.get(r[21], ""), r[22], r[23]])
    t = t.replace("__ENCARTE__", json.dumps(enc, ensure_ascii=False, separators=(",", ":")))
    os.makedirs(f"{RAIZ}/relatorio", exist_ok=True)
    open(f"{RAIZ}/relatorio/precos_mercados.html", "w", encoding="utf-8").write(t)
    return len(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--so", nargs="*", default=list(COLETORES))
    ap.add_argument("--data", default=datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-3))).date().isoformat())
    a = ap.parse_args()
    os.makedirs(TMP, exist_ok=True)
    with ThreadPoolExecutor(len(a.so)) as ex:
        res = dict(ex.map(coletar, a.so))
    for n, r in res.items(): print(n, r, flush=True)
    ok = [n for n, r in res.items() if r["ok"]]
    if ok:
        subprocess.run([sys.executable, f"{S}/base_historica/base.py", "carregar", a.data] + ok, check=True)
    subprocess.run([sys.executable, f"{S}/comparar_mercados.py"], check=True, capture_output=True)
    subprocess.run([sys.executable, f"{S}/encartes/comparar_encarte.py"], check=True, capture_output=True)
    n = pagina(a.data)
    resumo = dict(data=a.data, mercados=res, produtos_comparaveis=n)
    json.dump(resumo, open(f"{RAIZ}/relatorio/ultima_atualizacao.json", "w"), ensure_ascii=False, indent=1)
    shutil.rmtree(TMP, ignore_errors=True)
    print(json.dumps(resumo, ensure_ascii=False))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
