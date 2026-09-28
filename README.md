# Meu Caixa

Controle financeiro pessoal em uma página só: gastos do cartão com categoria e parcelas, fatura por cartão, proventos e quanto cada pessoa que usou seu cartão está te devendo.

## Como usar

Abra o `index.html` no navegador. Os dados ficam salvos só no navegador em que você usa a página (localStorage). Nenhum dado é enviado para servidor.

- **Gasto:** marca, data, valor total, categoria (Despesa fixa, Lazer, Alimentação, Outros), cartão, parcelas, fatura em que entra e, se for o caso, quem usou o cartão.
- **Provento:** salário e outras entradas.
- **Me pagaram:** pagamentos de quem te deve.
- **Cartões:** cadastro dos seus cartões (sem número, só nome, bandeira, vencimento e cor).

## Instalar no celular

O app é um PWA: abra https://thiagonascimento-cloud.github.io/meu-caixa/ no Chrome do Android e toque em **Instalar app** (ou **⋮ → Adicionar à tela inicial**). Funciona sem internet.

## Backup

Como os dados ficam no aparelho, use **Exportar backup** para salvar um arquivo `.json` e **Importar backup** para carregar em outro aparelho.

Política de privacidade: https://thiagonascimento-cloud.github.io/meu-caixa/privacidade.html

## Sincronizar com Planilha Google

Para ter os mesmos dados no iPhone, Android e computador, conecte uma Planilha Google sua. O código do Google está em [`google-apps-script/Code.gs`](google-apps-script/Code.gs); depois de implantar, cole o endereço e a chave no app, em **Sincronizar com Google**. A chave gerada fica só na sua conta Google e nos seus aparelhos.
