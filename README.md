# Bot-Automatizador-Planilhas

# Projeto de Organização de Dados de Planilha por Bairros

Este projeto em Python automatiza o processo de organização dos dados de uma planilha do Excel. Ele percorre a base de dados existente, identifica os bairros mencionados e, em seguida, cria uma aba separada para cada bairro, transferindo as informações correspondentes para facilitar o acesso e a leitura.

## Arquivos

- **codigo_automacao.py**: Contém o código Python que realiza a automação descrita.
- **Planilha_Original.xlsx**: A planilha Excel original, sem modificações.
- **Planilha_Pos_Automacao.xlsx**: A planilha Excel gerada após a execução do código, com os dados organizados por bairros em abas separadas.

## Funcionalidades

- **Criação Automática de Abas**: Para cada bairro encontrado na base de dados, uma nova aba é criada, caso ainda não exista.
- **Transferência de Dados**: As informações das pessoas e seus bairros são transferidas da aba principal para as respectivas abas dos bairros.
- **Preservação do Estilo**: O estilo das células (formatação) é mantido ao transferir os dados.

## Estrutura do Código

- **criar_aba(bairro, arquivo_bairros, estilos_cabecalho)**: Cria uma nova aba no arquivo do Excel para o bairro especificado, com os cabeçalhos de coluna definidos.
- **tranferir_informacoes_aba(aba_origem, aba_destino, linha_origem)**: Transfere as informações de uma linha específica da aba principal para a aba do bairro correspondente.
- **arquivo_bairros = load_workbook("Planilha_Original.xlsx")**: Carrega a planilha do Excel com os dados.
- **aba_basedados = arquivo_bairros["Base de Dados"]**: Define a aba principal que contém os dados de origem.
- **arquivo_bairros.save("Planilha_Pos_Automacao.xlsx")**: Salva o arquivo Excel atualizado com os dados organizados.

## Requisitos

- Python 3.x
- Biblioteca `openpyxl`

Para instalar a biblioteca `openpyxl`, utilize o seguinte comando:

```bash
pip install openpyxl

Como Usar
Certifique-se de ter a planilha Excel original como Planilha_Original.xlsx na mesma pasta que o script.
Execute o script codigo_automacao.py. Ele criará abas separadas para cada bairro encontrado e transferirá as informações.
O arquivo atualizado será salvo como Planilha_Pos_Automacao.xlsx na mesma pasta onde o script foi executado.
Exemplo de Planilha de Entrada
A aba "Base de Dados" da planilha original deve conter as seguintes colunas:

Data de Nascimento
Pessoa
Bairro
Cada linha subsequente deve conter os dados correspondentes para uma pessoa.

Contribuição
Se você deseja contribuir para este projeto, sinta-se à vontade para abrir uma issue ou enviar um pull request.

Licença
Este projeto está licenciado sob a Licença MIT. Consulte o arquivo LICENSE para obter mais informações.

