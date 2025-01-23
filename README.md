Data Cleaning Script

Este repositório contém um script em Python desenvolvido para consolidar, limpar e processar dados de arquivos Excel. Ele realiza operações de limpeza e organização, gerando relatórios detalhados e resumos de compras por cliente, incluindo cálculos do período médio de compras.

Funcionalidades:
Consolida arquivos Excel de múltiplos anos em um único DataFrame.
Realiza limpeza de dados, como:
Remoção de duplicatas.
Exclusão de registros irrelevantes ou inconsistentes.
Normalização de colunas (ex.: remoção de caracteres especiais).
Calcula o período médio entre compras para cada cliente.

Gera um arquivo Excel com:
Aba "Dados Detalhados": Registros processados.
Aba "Resumo": Consolidação das compras e estatísticas por cliente.
Estrutura do Arquivo de Saída
Aba "Dados Detalhados"
Contém todos os registros processados, incluindo colunas como:
CLIENTE_NOME: Nome do cliente.
CLIENTE_ID: Identificação do cliente.
CHASSI: Identificador único do veículo.
DATA_VENDA: Data da venda.
VALOR_VENDA: Valor da venda.
Aba "Resumo"

Apresenta um resumo consolidado com:
CLIENTE_NOME: Nome do cliente.
CLIENTE_ID: Identificação do cliente.
CIDADE_CLIENTE: Cidade do cliente.
QUANTIDADE_COMPRAS: Número total de compras realizadas.
PERIODO_MEDIO_MES: Período médio entre compras em meses.

Como Configurar:
Certifique-se de que os arquivos Excel estão localizados no diretório configurado no script:
base_dir = "C:/automacao mkt/base de dados"
Defina o caminho onde o arquivo de saída será salvo:
arquivo_saida = "C:/automacao mkt/base de dados/quantidade_compras_por_cliente.xlsx"

Como Executar:
Clone o repositório:
git clone https://github.com/seu-usuario/data-cleaning-script.git
cd data-cleaning-script
Instale as dependências necessárias:
pip install pandas openpyxl
Execute o script:

python main.py

O arquivo Excel consolidado estará disponível no caminho configurado em arquivo_saida.
Observações Importantes
O script ignora arquivos temporários ou corrompidos que começam com ~$.

Dados sensíveis, como o nome do cliente e identificadores, são processados e podem ser censurados ou mascarados conforme necessário.
Certifique-se de que o diretório de entrada contenha apenas arquivos relevantes.

Contribuindo

Contribuições são bem-vindas! Para sugerir melhorias ou relatar problemas, envie uma solicitação via "Issues" ou abra um pull request.

Autor: Gabriel Matuck
Contato: gabriel.matuck1@gmail.com
