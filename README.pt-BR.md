# Excel File Merger

[English](README.md) | **Português**

## Visão geral

O Excel File Merger é um aplicativo de desktop que mescla as planilhas de
vários arquivos `.xlsx` em um único arquivo.

## Funcionalidades

- **Seleção de arquivos ou pasta** — adicione arquivos `.xlsx` individuais ou
  uma pasta inteira pelo explorador de arquivos.
- **Mesclagem por nome de planilha** — planilhas com o mesmo nome em
  diferentes arquivos são combinadas em uma só; a ordem segue o primeiro
  arquivo e nomes extras são anexados ao final.
- **Preservação de cor de fundo** — as cores de fundo RGB das células são
  mantidas durante a mesclagem e permanecem alinhadas aos seus dados mesmo
  após edições no schema.
- **Schemas editáveis** — o aplicativo pré-gera um schema por planilha
  mesclada; renomeie, reordene, adicione ou remova colunas no editor
  integrado antes de salvar.
- **Validação flexível** — inconsistências de dados (planilha presente em
  apenas alguns arquivos, cabeçalhos divergentes, vazios removidos) aparecem
  como avisos em um painel dedicado; a mesclagem nunca é bloqueada por eles.

## Como funciona

1. **Selecione** — adicione arquivos `.xlsx` ou uma pasta.
2. **Mescle** — as planilhas são pareadas por nome e suas linhas combinadas;
   linhas/colunas em branco são removidas e as cores das células capturadas.
3. **Revise os schemas** — uma aba por planilha mesclada permite renomear,
   reordenar, adicionar ou remover colunas.
4. **Salve** — os schemas editados são aplicados e o arquivo mesclado é
   gravado no `.xlsx` de saída que você escolher.

## Pré-requisitos

- Python 3.14+
- [uv](https://docs.astral.sh/uv/)
- [just](https://github.com/casey/just)

## Build e execução

```bash
# instalar dependências
uv sync

# buildar o executável (saída: dist/)
just build

# executar a partir do código (desenvolvimento)
uv run python src/gui_main.py
```
