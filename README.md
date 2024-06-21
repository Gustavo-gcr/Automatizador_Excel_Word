# Automatizador_Excel_Word

Este sistema é desenvolvido em Python para automatizar a geração de contratos personalizados em formato Word (.docx), utilizando dados extraídos de uma planilha Excel. A seguir, um resumo detalhado do funcionamento do código:

1. **Importação de Bibliotecas**:
   - `from docx import Document`: Importa a biblioteca necessária para manipular documentos Word.
   - `from datetime import datetime`: Importa a biblioteca para manipular datas.
   - `import pandas as pd`: Importa a biblioteca Pandas para manipulação de dados tabulares.

2. **Leitura dos Dados**:
   - `tabela = pd.read_excel('Pasta1.xlsx')`: Carrega os dados da planilha Excel denominada 'Pasta1.xlsx' em um DataFrame do Pandas.

3. **Iteração pelas Linhas da Tabela**:
   - `for linha in tabela.index`: Itera sobre cada linha do DataFrame.

4. **Carregamento do Modelo de Contrato**:
   - `documento = Document("Contrato.docx")`: Carrega o documento Word 'Contrato.docx' como modelo para a geração dos contratos personalizados.

5. **Extração das Informações da Planilha**:
   - `nome = tabela.loc[linha, 'Nome']`: Extrai o nome do cliente.
   - `item1 = tabela.loc[linha, 'Item1']`: Extrai o primeiro item.
   - `item2 = tabela.loc[linha, 'Item2']`: Extrai o segundo item.
   - `item3 = tabela.loc[linha, 'Item3']`: Extrai o terceiro item.

6. **Criação de um Dicionário de Referências**:
   - `referencias = { ... }`: Cria um dicionário que mapeia os códigos de substituição no documento Word para os valores extraídos da planilha Excel e a data atual.

7. **Substituição de Texto no Documento**:
   - `for paragrafo in documento.paragraphs`: Itera sobre cada parágrafo do documento.
   - `for cod in referencias`: Para cada código no dicionário de referências, substitui o texto correspondente no parágrafo pelo valor mapeado.

8. **Salvamento do Documento Personalizado**:
   - `documento.save(f"Contrato-{nome}.docx")`: Salva o documento Word com o nome do cliente incluído no nome do arquivo.

### Observação

Este sistema automatiza a geração de contratos personalizados em formato Word (.docx) a partir de dados extraídos de uma planilha Excel. Não há integração com banco de dados neste projeto.

### Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues e pull requests.

### Licença

Este projeto está licenciado sob a Licença MIT. Consulte o arquivo LICENSE para obter mais informações.

## Autores

- [@Gustavo-gcr](https://github.com/Gustavo-gcr)
