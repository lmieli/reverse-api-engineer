# Extração de mercados: convenções

Vale para a Bramil, o Superpão e qualquer mercado novo que entrar na comparação.

1. **Sempre extrair o link da imagem do produto**, junto com nome, preço, departamento e unidade.
   Guardar a URL completa (monte o prefixo do CDN quando a API devolver só o nome do arquivo) e
   deixar vazio quando o site não tiver foto.
   - Bramil: `https://produto-assets-vipcommerce-com-br.br-se1.magaluobjects.com/250x250/<imagem>`
   - Superpão: campo `image` do produto (CDN `phygital-files.mercafacil.com`).
2. **De-para 1 para 1.** Cada produto de um mercado tem no máximo um par no outro. Nunca 1 para N.
   Produto sem equivalente fica "Sem par". Veja `compare_prices.py` (atribuição global, tamanho e
   embalagem iguais, variações diferentes não pareiam).
3. Preço em promoção: guardar o preço normal e o promocional separados.
4. Usar só o que o site já entrega ao visitante. Se a API recusar o acesso direto (403), não
   contornar: ler as páginas públicas permitidas pelo robots.txt.
