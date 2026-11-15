"""Captura como o site do Tio Dongo entrega os produtos, para eu escrever o coletor.

Rode em uma rede brasileira (onde https://www.tiodongo.com.br/ abre no navegador).
Visita a página inicial, o robots.txt, o sitemap e algumas categorias; salva o HTML e as chamadas JSON
que o site faz. Não faz login e só lê páginas públicas. Gera a pasta captura_tiodongo/ e o arquivo
captura_tiodongo.zip, que você me envia.

Uso:
  pip install playwright && playwright install chromium
  python3 capturar.py
"""
import json, os, re, shutil, sys, time
from playwright.sync_api import sync_playwright

BASE = "https://www.tiodongo.com.br"
OUT = "captura_tiodongo"
os.makedirs(OUT, exist_ok=True)
chamadas = []


def salvar(nome, conteudo):
    with open(os.path.join(OUT, nome), "w", encoding="utf-8") as f:
        f.write(conteudo)


def main():
    with sync_playwright() as p:
        nav = p.chromium.launch()
        ctx = nav.new_context(locale="pt-BR", viewport={"width": 1366, "height": 900})
        pag = ctx.new_page()

        def resp(r):
            try:
                tipo = r.headers.get("content-type", "")
                if r.request.resource_type in ("xhr", "fetch") or "json" in tipo:
                    corpo = r.text()[:6000]
                    chamadas.append({"url": r.url, "status": r.status, "metodo": r.request.method, "tipo": tipo,
                                     "corpo_inicio": corpo})
            except Exception:
                pass
        pag.on("response", resp)

        r = pag.goto(BASE + "/", wait_until="networkidle", timeout=60000)
        print("página inicial:", r.status if r else "sem resposta")
        if not r or r.status >= 400:
            print("O site não abriu. Confirme que você está em uma rede brasileira e tente de novo.")
        pag.wait_for_timeout(2000)
        salvar("home.html", pag.content())
        pag.screenshot(path=os.path.join(OUT, "home.png"))

        for nome in ("robots.txt", "sitemap.xml"):
            try:
                x = ctx.request.get(f"{BASE}/{nome}", timeout=30000)
                print(nome, x.status)
                salvar(nome, x.text()[:200000])
            except Exception as e:
                print(nome, "erro", e)

        # abre até 3 links que parecem categoria e rola cada um para disparar a lista de produtos
        links = pag.eval_on_selector_all("a[href]", "els => els.map(e => e.href)")
        cand = []
        for l in links:
            if l.startswith(BASE) and re.search(r"/(categoria|departamento|produtos|c/|collections|catalogo)", l) and l not in cand:
                cand.append(l)
        for i, url in enumerate(cand[:3]):
            try:
                pag.goto(url, wait_until="networkidle", timeout=60000)
                for _ in range(6):
                    pag.mouse.wheel(0, 2500); pag.wait_for_timeout(700)
                salvar(f"categoria_{i+1}.html", pag.content())
                print("categoria", i + 1, url)
            except Exception as e:
                print("categoria erro", url, e)
        salvar("todos_os_links.txt", "\n".join(sorted(set(links))))
        nav.close()

    salvar("chamadas_json.json", json.dumps(chamadas, ensure_ascii=False, indent=1))
    shutil.make_archive(OUT, "zip", OUT)
    print(f"\nPronto: {OUT}.zip ({len(chamadas)} chamadas JSON capturadas). Envie esse arquivo.")


if __name__ == "__main__":
    sys.exit(main())
