# Invoice Slab Filler

Chrome extension que carrega o Excel exportado pelo **PO Packing List Tool** e gera TSV por bundle para colar na modal *Packinglist Details* do Stone Profits.

A extensão **não abre sozinha** — só quando você clicar no ícone dela na barra do Chrome.

## Como usar

1. Gera o Excel no [PO Packing List Tool](https://marbleexpresslv.github.io/PO-PACKING-LIST-TOOL/) (botão "Export Excel").
2. Abre a página da invoice no Stone Profits e clica no ícone azul (Update Slab Info) do produto pra abrir a modal.
3. Na modal, garante que o toggle de unidade está em **m**.
4. Clica no ícone da extensão na barra do Chrome → faz upload do Excel.
5. O popup mostra a lista de bundles. Clica em **Copy TSV** do bundle correto.
6. Volta na modal, clica na primeira célula vazia (Length da linha 1) e faz **Cmd+V**.
7. Confere e clica em **Add Slabs**.

O Excel fica salvo na sessão do Chrome — você só precisa carregar uma vez. Pra trocar, clica em "Carregar outro Excel" no popup.

## Instalação (modo desenvolvedor)

1. Abre `chrome://extensions/`
2. Ativa **Modo do desenvolvedor** (canto superior direito)
3. Clica em **Carregar sem compactação**
4. Seleciona a pasta `extension/`

## Estrutura esperada do Excel

Gerado automaticamente pelo PO Packing List Tool:

- **Aba `Summary`**: Supplier, Packing List, PO Number, Container, Unit
- **Abas `BUNDLE <numero>`**: uma por bundle, colunas `Length, Width, Quantity, Block, Bundle, Supplier Reference`

Se a unidade do Excel ≠ "m", o popup mostra aviso amarelo.

## Arquivos

- `manifest.json` — MV3, popup
- `popup.html` + `popup.js` — UI e lógica
- `lib/xlsx.full.min.js` — SheetJS (parser de Excel, local)
