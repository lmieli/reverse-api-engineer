# Extração de mercados: convenções

Vale para a Bramil, o Superpão e qualquer mercado novo que entrar na comparação.

1. **Sempre extrair o link da imagem do produto**, junto com nome, preço, departamento e unidade.
   Guardar a URL completa (monte o prefixo do CDN quando a API devolver só o nome do arquivo) e
   deixar vazio quando o site não tiver foto.
   - Bramil: `https://produto-assets-vipcommerce-com-br.br-se1.magaluobjects.com/250x250/<imagem>`
   - Superpão: campo `image` do produto (CDN `phygital-files.mercafacil.com`).
2. **De-para 1 para 1.** Cada produto de um mercado tem no máximo um par no outro. Nunca 1 para N.
   Produto sem equivalente fica "Sem par".
   - A fonte da verdade é `depara_bramil_superpao.csv` (SKU Bramil -> id Superpão). Linhas com
     `origem = revisado_sem_par` registram produtos conferidos que não têm equivalente.
   - `compare_prices.py` aplica o CSV primeiro; a heurística (nome + tamanho + embalagem) só sugere
     pares para produtos novos, marcados "Novo (sugestão, revisar)" na planilha.
   - A cada atualização: rode a comparação, revise as sugestões novas e acrescente as decisões ao CSV.
     Mercado novo = um CSV de de-para novo, no mesmo formato.
   - Preço por kg: quando um lado vende por quilo (unidade KG sem peso no nome) e o outro informa o
     peso, a comparação é feita em R$/kg.
   - Peso diferente no cadastro: o mesmo produto (marca, linha e sabor) às vezes aparece com peso
     diferente nos dois sites (ex.: 90 g x 80 g, embalagem reduzida). Esses pares são conferidos à mão,
     gravados com `origem = revisado_peso` e comparados em R$ por kg ou litro.
3. Preço em promoção: guardar o preço normal e o promocional separados.
4. Usar só o que o site já entrega ao visitante. Se a API recusar o acesso direto (403), não
   contornar: ler as páginas públicas permitidas pelo robots.txt.

## Base histórica (`scripts/base_historica/`)
- `base.py carregar [AAAA-MM-DD]` lê os arquivos atuais de cada mercado e grava um snapshot em
  `precos_historico.db` (SQLite): uma linha de preço por produto por data (`preco`), cadastro do produto
  (`produto`, com primeira/última coleta, imagem e URL), controle de coletas (`coleta`) e o de-para.
- Para ter histórico, rode a coleta de cada mercado e depois `base.py carregar` a cada data. Reexecutar no
  mesmo dia substitui aquele dia. Nada é apagado: produto que some do site mantém o histórico.
- Mercados: `bramil` (API VipCommerce), `superpao` e `remercado` (Mercafacil, páginas públicas do sitemap,
  `scripts/remercado_scraper/scraper.py`). `tiodongo` está cadastrado, mas o site devolve 403 para IPs fora
  do Brasil; o coletor depende de rodar em uma rede onde o site abre. Não contornar o bloqueio.
- Remercado ainda não tem de-para com a Bramil/Superpão; fazer no mesmo formato (CSV 1 para 1).

## De-para do Remercado (`scripts/depara_bramil_remercado.csv`)
- Mesmo formato do Bramil x Superpão (1 para 1; `bramil_sku`, `remercado_id`, `origem`, `nota`), com 2.219 pares.
  `origem = auto` (nota alta, todas as palavras casam) ou `revisado` (conferido item a item).
- Remercado sem par: pares com nota < 93 e candidatos com nota >= 50 foram revisados um a um; os candidatos
  com nota 40-49 (cerca de 500) ainda não foram revisados item a item.
- Hortifruti e carnes vendidos por kg entram quando o produto é o mesmo (ex.: banana ouro kg x bandeja 500 g);
  o preço comparável por kg deve ser usado ao comparar.
- Os dados também ficam na tabela `depara_remercado` da base histórica.

## De-para Remercado x Superpão (`scripts/depara_remercado_superpao.csv`)
- Colunas: `remercado_id`, `superpao_id`, `origem` (`auto` ou `revisado`), `nota`. 1 para 1, 1.449 pares.
- Pares com nota < 85 e candidatos com nota >= 62 entre os sem par foram revisados item a item; candidatos
  com nota 55-61 (cerca de 400) não foram revisados.
- Vira a tabela `depara_remercado_superpao` na base histórica (`base.py carregar`).

## Comparação entre todos os mercados (`scripts/comparar_mercados.py`)
- `python3 scripts/comparar_mercados.py` junta os três de-paras em grupos (um produto = no máximo um item por
  mercado) e gera `comparacao_precos_todos_mercados.xlsx` (abas Comparação e Resumo) e `scripts/produto_grupo.csv`.
- Quando os de-paras se contradizem em um grupo, mantém o maior conjunto de pares sem contradição.
- Itens vendidos por kg contra embalagem com peso: compara em R$/kg. Item sem peso (ex.: bandeja) fica de fora.
- Linhas com diferença > 100% saem marcadas "conferir".
- Novo mercado: criar o de-para 1:1 dele e acrescentar a aresta em `ler_arestas`.

## Tio Dongo (`scripts/tiodongo/`)
- Plataforma VipCommerce (a API vive em services.vipcommerce.com.br, não no domínio do site, e responde de fora
  do Brasil). `node scraper.mjs tiodongo_produtos.json` coleta tudo em ~4 min. API.md documenta as chamadas.
- Preço: `preco` é o preço normal. Só a oferta "Encarte" entra como promoção; "Preco Fidelidade" é preço só do
  clube e não é usado como preço de comparação.
- De-para Bramil x Tio Dongo por código de barras (`scripts/depara_bramil_tiodongo.csv`, origem `ean`). O EAN
  do catálogo nem sempre bate com o produto: pares com nomes muito distantes ou tamanho diferente foram removidos.
- Tio Dongo x Superpão/Remercado: entram no comparador por transitividade, via Bramil. Os 2.9 mil produtos do
  Tio Dongo sem EAN em comum com a Bramil ainda não têm de-para com Superpão/Remercado.
- Carregar só o Tio Dongo na base: `python3 scripts/base_historica/base.py carregar AAAA-MM-DD tiodongo`.

### Tio Dongo x Superpão / Remercado
- `scripts/depara_tiodongo_superpao.csv` (625 pares) e `scripts/depara_tiodongo_remercado.csv` (679 pares), por nome,
  1 para 1, só para os produtos do Tio Dongo sem par por EAN na Bramil. Revisados item a item; também ficam na
  tabela `depara_tiodongo_mercado` da base.
- Produtos da padaria própria do Tio Dongo (marcados "Tio Dongo" no nome) não são pareados com os de outros mercados.
- Itens por kg do Tio Dongo: `preco` é o preço da porção estimada; o preço por kg está em `preco_original`. A oferta
  "Encarte" também vem pela porção e é convertida para R$/kg na mesma proporção. O comparador e a base usam R$/kg.
  Sem `preco_original`, estima pelo peso da porção no nome.

### Bramil x Tio Dongo por nome
- Além do EAN, `depara_bramil_tiodongo.csv` tem pares por nome (`origem` = `nome` ou `nome_auto`), revisados item a item,
  para produtos do Tio Dongo sem EAN em comum com a Bramil. Total: 2.990 pares (2.709 por EAN, 281 por nome).

### Revisão das linhas "conferir" (diferença > 100%)
- Revisadas uma a uma. Pares errados (pack x unidade, sabor/modelo diferente, sachê x vidro, cadastro errado do site)
  foram desfeitos nos de-paras. As que restam são o mesmo produto com preço de fato diferente (promoção, bandeja x kg).

## Atualização diária (`scripts/atualizar_tudo.py`)
- `python3 scripts/atualizar_tudo.py` coleta os 4 mercados em paralelo (~45 min, a Bramil é a mais lenta), grava o dia
  na base histórica, refaz `comparacao_precos_todos_mercados.xlsx` e gera `relatorio/precos_mercados.html` (a página)
  e `relatorio/ultima_atualizacao.json` (resumo). Se um mercado falhar ou vier com menos de 80% dos produtos da coleta
  anterior, mantém os dados antigos e não grava esse mercado no dia.
- Os de-paras não são refeitos: produtos novos entram na base, mas só aparecem na comparação depois de pareados.
- Roda todo dia por volta das 4h (Brasília) por uma tarefa agendada, que também republica a página e faz o commit.
