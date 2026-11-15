# Tio Dongo — mapa das APIs

Site: https://www.tiodongo.com.br (plataforma **VipCommerce**).
Os dados vêm de `https://services.vipcommerce.com.br/api-admin/v1`, não do domínio do site.
Mapeado em 2026-11-15 pelo Chrome DevTools; o que está marcado "verificado" foi chamado de fato pelo `scraper.mjs`.

## Constantes da loja

| Item | Valor |
|---|---|
| Domínio (header `domainkey`) | `tiodongo.com.br` |
| Organização (header `organizationid`) | `422` |
| Filial | `1` |
| Centro de distribuição | `1` (retirada na loja, Rua Clarindo da Rosa Teixeira) |
| Filial de aplicativos | `472` |

Headers em toda chamada: `domainkey`, `organizationid`, `accept: application/json`, `content-type: application/json`,
`origin` e `referer` do site, `authorization: Bearer <token>` (exceto no login).
O site também manda `sessao-id`, mas o scraper funciona sem ele.

## O que chamar para atualizar os preços (verificado)

Ordem usada pelo `scraper.mjs`. Uma atualização completa são cerca de 280 chamadas, em ~2 min com pausa de 0,3 s.

1. **Login anônimo da loja** (um token por execução)
   `POST /org/422/auth/loja/login`
   Corpo: `{"domain":"tiodongo.com.br","username":"loja","key":"<chave pública do site>"}`
   Resposta: `{"success":true,"data":"<JWT>"}`. A chave está no `scraper.mjs`; é a que o site embute para todo visitante.
   Se o login passar a falhar com 401/403, abra o site no navegador e copie a chave nova da requisição de login.

2. **Árvore de departamentos** (lista os ids a percorrer)
   `GET /org/422/filial/1/centro_distribuicao/1/loja/classificacoes_mercadologicas/departamentos/arvore`
   `data[]` = departamentos (nível `Departamento`) com `children[]` = seções. Campos: `classificacao_mercadologica_id`,
   `descricao`, `link`, `total_ofertas`.
   Departamentos (ids de hoje): ver o resultado de `departamentos/arvore`; o Açougue é o `7`.

3. **Produtos de cada departamento** (paginado, 20 por página)
   `GET /org/422/filial/1/centro_distribuicao/1/loja/classificacoes_mercadologicas/departamentos/{id}/produtos?page={n}`
   `paginator` traz `total_pages` e `total_items`. `data[]` traz os produtos (a seção vem em `secao_id`).
   Percorra `page` de 1 até `total_pages` para cada departamento da árvore.

### Campos de produto que importam

| Campo da API | Significado |
|---|---|
| `produto_id` | id estável do produto na loja |
| `descricao` | nome |
| `preco` | preço normal atual (string) |
| `preco_original` + `exibe_preco_original` | preço por unidade de venda (kg etc.) quando exibido |
| `em_oferta`, `oferta` | oferta ativa; em `oferta.preco_oferta` está o preço promocional |
| `oferta.nome` / `oferta.tag` | tipo da oferta. Maioria hoje é `Preco Fidelidade` (preço só para o clube) |
| `disponivel` | em estoque ou não |
| `unidade_sigla` | `UN`, `KG`... |
| `codigo_barras` | EAN, usado para casar com os outros mercados |
| `marca` / `marca_id` | vazios em todos os 5.622 produtos |
| `imagem` | só o nome do arquivo; a URL base não foi confirmada |

## Outras APIs vistas, ainda não usadas

Observadas no tráfego do site (status 200), sem uso no coletor:

- Vitrines da home: `GET .../loja/vitrines/produtos?vitrine_ids={uuid}&page=1&limit=10`
- Filtros de um departamento: `GET .../classificacoes_mercadologicas/departamentos/{id}/produtos/filtros`
- Banners: `GET .../loja/banner/localizacao/classificacao-mercadologica/dispositivo/S?cm.id={id}`
- Configuração da loja: `GET /org/422/build/app/configuracao/1`, `GET /org/422/loja/centros_distribuicoes/1`,
  `GET /organizacoes/filiais/dominio/tiodongo.com.br`

Rotas do site (front-end), úteis para montar o link do produto:
`/produtos/departamento/:departamento`, `/produtos/departamento/:departamento/:secao`,
`/produtos/detalhe/:id/:slug`, `/produtos/colecoes/:id/:slug`, `/produtos/ofertas/tags/:tag`,
`/produtos/mais-vendidos/:departamento`, `/busca`.
Não foram testadas: busca por texto, detalhe de um produto isolado e troca de filial/CD (preço pode variar por filial).

## Cuidados

- O domínio `www.tiodongo.com.br` devolve 403 a clientes que não parecem navegador (curl/PowerShell puros). A API
  `services.vipcommerce.com.br` respondeu normalmente a `fetch` do Node.
- Não testado de fora do Brasil nem de IP de datacenter.
- Só leitura de catálogo público, sem conta de cliente. Mantenha a pausa entre páginas.
- Só existe a filial `1`/CD `1`; preços e disponibilidade são dessa loja.

## Rodar

```
node scraper.mjs tiodongo_produtos.json
```
