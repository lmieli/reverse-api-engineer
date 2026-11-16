"""Cliente para a loja Bramil em Casa (https://www.bramilemcasa.com.br/).

Plataforma VipCommerce. Fluxo: sessão -> login anônimo (usuário público "loja",
chave embutida no front-end) -> token Bearer -> listagem de produtos.
Preços variam por centro de distribuição (loja física).
"""
import json
import time
import urllib.request

BASE = "https://services-beta.vipcommerce.com.br/api-admin/v1"
DOMAIN = "bramilemcasa.com.br"
ORG = 53
FILIAL = 1
# Chave pública enviada pelo próprio site no login anônimo.
PUBLIC_KEY = "df072f85df9bf7dd71b6811c34bdbaa4f219d98775b56cff9dfa5f8ca1bf8469"
DEFAULT_CD = 17  # Bramil PV10 - Nova Friburgo (padrão do site)


class BramilClient:
    def __init__(self, centro_distribuicao=DEFAULT_CD):
        self.cd = centro_distribuicao
        self.token = ""
        self.session_id = ""
        self._login()

    def _req(self, method, path, body=None):
        headers = {
            "accept": "application/json",
            "content-type": "application/json",
            "domainkey": DOMAIN,
            "organizationid": str(ORG),
            "authorization": f"Bearer {self.token}".strip(),
            "sessao-id": self.session_id,
            "user-agent": "Mozilla/5.0",
        }
        data = json.dumps(body).encode() if body is not None else None
        req = urllib.request.Request(BASE + path, data=data, method=method, headers=headers)
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.load(r)

    def _login(self):
        out = self._req("POST", f"/org/{ORG}/auth/loja/login",
                        {"domain": DOMAIN, "username": "loja", "key": PUBLIC_KEY})
        self.token = out["data"]
        self.session_id = self._req("GET", f"/org/{ORG}/loja/sessao_cliente")["data"]["sessao_id"]

    def _loja(self):
        return f"/org/{ORG}/filial/{FILIAL}/centro_distribuicao/{self.cd}/loja"

    def departamentos(self):
        """Árvore de departamentos (id, descricao, children)."""
        return self._req("GET", self._loja() + "/classificacoes_mercadologicas/departamentos/arvore")["data"]

    def produtos_departamento(self, departamento_id, por_pagina=200, pausa=0.3):
        """Itera todos os produtos de um departamento (paginação automática, `por_pagina` itens por requisição)."""
        page = 1
        while True:
            out = self._req("GET", self._loja() + f"/classificacoes_mercadologicas/departamentos/{departamento_id}"
                            f"/produtos?page={page}&limit={por_pagina}")
            items = out.get("data") or []
            if not items:
                return
            yield from items
            total = (out.get("paginator") or {}).get("total_pages")
            if total and page >= total:
                return
            page += 1
            time.sleep(pausa)
