# Excel Merger

O script junta arquivos excel com tabelas que seguem uma estrutura padrão.

### ❓ Como funciona
O script recebe uma pasta com arquivos excel e junta todos eles em um único arquivo.

### 🥚 Pré-requisitos
* Python 3.14+
* [uv](https://docs.astral.sh/uv/)
* [just](https://github.com/casey/just)

### ⚙️ Instalação
#### Windows
```bash
uv sync
just build
```
Após buildar o projeto com [pyinstaller](https://pyinstaller.org/) por meio do comando 'just build', o executável e suas dependências estarão dentro da pasta 'dist'.

### 🚀 Uso

1. Forma mais simples de usar a ferramenta. Um arquivo 'merged.pdf' será criado na pasta em que a ferramenta foi utilizada.
```bash
excelmerger ./pasta 
```

2. Mesclar PDFs e especificar o local de escrita e o nome do arquivo PDF.
```bash
excelmerger ./pasta --o ./salvar/aqui/excel.xlsx
```

3. Use a _flag_ 'overwrite' para sobrescrever mesclagens com o mesmo nome.
```bash
excelmerger ./pasta --o ./salvar/aqui/excel.xlsx --overwrite
```
