// Extrai produtos e preços públicos do Tio Dongo (https://www.tiodongo.com.br).
//
// Plataforma VipCommerce. O próprio site faz um login anônimo de loja (usuário "loja" e a chave que vem
// no código do site) e depois lista os produtos por departamento. Este script repete exatamente esse fluxo,
// sem conta de cliente: só leitura de catálogo público.
//
// Uso (Node 18+):  node scraper.mjs [saida.json]
import { writeFileSync } from "node:fs";

const SITE = "https://www.tiodongo.com.br";
const API = "https://services.vipcommerce.com.br/api-admin/v1";
const ORG = 422, FILIAL = 1, CD = 1;
const DOMINIO = "tiodongo.com.br";
const CHAVE_LOJA = "df072f85df9bf7dd71b6811c34bdbaa4f219d98775b56cff9dfa5f8ca1bf8469"; // pública, vem no site
const PAUSA_MS = 300;

const base = {
  "content-type": "application/json",
  accept: "application/json",
  domainkey: DOMINIO,
  organizationid: String(ORG),
  origin: SITE,
  referer: SITE + "/",
  "user-agent": "Mozilla/5.0",
};
const dorme = (ms) => new Promise((r) => setTimeout(r, ms));

async function req(url, opts = {}, tentativas = 3) {
  for (let i = 1; i <= tentativas; i++) {
    try {
      const r = await fetch(url, opts);
      if (r.ok) return await r.json();
      if (r.status === 403 || r.status === 401) throw new Error(`HTTP ${r.status} em ${url}`);
    } catch (e) {
      if (i === tentativas || /HTTP 40[13]/.test(String(e))) throw e;
    }
    await dorme(2000 * i);
  }
}

async function main(saida) {
  const login = await req(`${API}/org/${ORG}/auth/loja/login`, {
    method: "POST",
    headers: base,
    body: JSON.stringify({ domain: DOMINIO, username: "loja", key: CHAVE_LOJA }),
  });
  const headers = { ...base, authorization: `Bearer ${login.data}` };
  const loja = `${API}/org/${ORG}/filial/${FILIAL}/centro_distribuicao/${CD}/loja`;

  const arvore = (await req(`${loja}/classificacoes_mercadologicas/departamentos/arvore`, { headers })).data;
  const nomes = new Map();
  const mapa = (n) => { nomes.set(n.classificacao_mercadologica_id, n); (n.children || []).forEach(mapa); };
  arvore.forEach(mapa);

  const vistos = new Set(), linhas = [];
  for (const dep of arvore) {
    let pagina = 1, total = 1;
    while (pagina <= total) {
      const j = await req(`${loja}/classificacoes_mercadologicas/departamentos/${dep.classificacao_mercadologica_id}/produtos?page=${pagina}`, { headers });
      total = j.paginator?.total_pages || 1;
      for (const p of j.data || []) {
        if (vistos.has(p.produto_id)) continue;
        vistos.add(p.produto_id);
        const secao = nomes.get(p.secao_id);
        linhas.push({
          id: String(p.produto_id),
          nome: p.descricao,
          marca: p.marca?.descricao ?? p.marca?.nome ?? p.marca ?? null,
          departamento: dep.descricao,
          categoria: secao?.link?.replace(/^\//, "") ?? null,
          preco: Number(p.preco),
          preco_original: p.exibe_preco_original ? Number(p.preco_original) : null,
          em_oferta: !!p.em_oferta,
          oferta: p.oferta ?? null,
          unidade: p.unidade_sigla ?? null,
          disponivel: !!p.disponivel,
          codigo_barras: p.codigo_barras ?? null,
          sku: p.sku ?? null,
          imagem_arquivo: p.imagem ?? null,
          url: `${SITE}/produtos/detalhe/${p.produto_id}/${p.link}`,
        });
      }
      console.log(`${dep.descricao} pág ${pagina}/${total} — ${linhas.length} produtos`);
      pagina++;
      await dorme(PAUSA_MS);
    }
  }
  writeFileSync(saida, JSON.stringify(linhas, null, 1), "utf8");
  console.log(`\nPronto: ${linhas.length} produtos em ${saida}`);
}

main(process.argv[2] || "tiodongo_produtos.json").catch((e) => { console.error("Erro:", e.message); process.exit(1); });
