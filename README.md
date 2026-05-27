# Excel Merger

O script junta arquivos excel com tabelas que seguem uma estrutura padrão.

### ❓ Como funciona
O script recebe uma pasta com arquivos excel e junta todos eles em um único arquivo. Opcionalmente,
o script também pode receber um arquivo .yaml para realizar formatação.

#### Exemplo de Formatação
```yaml
header:
  bg-color: "#ff7d99d1"
  text-color: "#ff333333"
  font-name: "Calibri"
  font-size: "14"
  text-bold: false
  text-italic: false
  text-strikethrough: false
  border: "thick #fff799d1"

cell:
  bg-color: "#ffffffff"
  text-color: "#ff333333"
  font-name: "Calibri"
  font-size: "11"
  text-bold: false
  text-italic: false
  text-strikethrough: false
  border: "mediumDashed #ffd17417"
```

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

1. Forma mais simples de usar a ferramenta. Um arquivo 'merged.xlsx' será criado na pasta em que a ferramenta foi utilizada.
```bash
excelmerger ./pasta 
```

2. Mesclar excel e especificar o local de escrita e o nome do arquivo excel.
```bash
excelmerger ./pasta --o ./salvar/aqui/excel.xlsx
```

3. Use a _flag_ 'overwrite' para sobrescrever mesclagens com o mesmo nome.
```bash
excelmerger ./pasta --o ./salvar/aqui/excel.xlsx --overwrite
```

4. Use a _flag_ 'style' para aplicar formatação nas tabelas.
```bash
excelmerger ./pasta --style formatting.yaml --o ./salvar/aqui/excel.xlsx --overwrite
```
