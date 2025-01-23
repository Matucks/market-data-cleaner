# Data Cleaning Script

Este repositório contém um script em Python projetado para consolidar, limpar e processar dados de arquivos Excel. Ele oferece funcionalidades de organização, análise e relatórios detalhados, incluindo o cálculo do período médio de compras por cliente.

---

## Funcionalidades

- **Consolidação de dados:**
  - Une dados de múltiplos arquivos Excel em um único DataFrame.
- **Limpeza de dados:**
  - Remove duplicatas.
  - Exclui registros irrelevantes ou inconsistentes.
  - Normaliza colunas (ex.: remoção de caracteres especiais).
- **Análise de dados:**
  - Calcula o período médio entre compras para cada cliente.

- **Geração de relatórios em Excel:**
  - **Aba "Dados Detalhados":** Registros processados com informações completas.
  - **Aba "Resumo":** Consolidação de compras e estatísticas por cliente.

---

## Estrutura do Arquivo de Saída

### Aba "Dados Detalhados"
Contém todos os registros processados, incluindo:
- **CLIENTE_NOME:** Nome do cliente.
- **CLIENTE_ID:** Identificação do cliente.
- **CHASSI:** Identificador único do veículo.
- **DATA_VENDA:** Data da venda.
- **VALOR_VENDA:** Valor da venda.

### Aba "Resumo"
Apresenta um resumo consolidado com:
- **CLIENTE_NOME:** Nome do cliente.
- **CLIENTE_ID:** Identificação do cliente.
- **CIDADE_CLIENTE:** Cidade do cliente.
- **QUANTIDADE_COMPRAS:** Número total de compras realizadas.
- **PERIODO_MEDIO_MES:** Período médio entre compras em meses.

---

## Configuração

1. Certifique-se de que os arquivos Excel estão localizados no diretório configurado no script:
   ```python
   base_dir = "C:/automacao mkt/base de dados"
   ```

2. Defina o caminho onde o arquivo de saída será salvo:
   ```python
   arquivo_saida = "C:/automacao mkt/base de dados/quantidade_compras_por_cliente.xlsx"
   ```

---

## Como Executar

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/data-cleaning-script.git
   cd data-cleaning-script
   ```

2. Instale as dependências:
   ```bash
   pip install pandas openpyxl
   ```

3. Execute o script:
   ```bash
   python main.py
   ```

4. O arquivo Excel consolidado estará disponível no caminho configurado em `arquivo_saida`.

---

## Observações Importantes

- O script ignora arquivos temporários ou corrompidos que começam com `~$`.
- Dados sensíveis, como nomes e identificadores de clientes, podem ser censurados ou mascarados conforme necessário.
- Certifique-se de que o diretório de entrada contenha apenas arquivos relevantes para o processamento.

---

## Contribuições

Contribuições são bem-vindas! Caso tenha sugestões de melhorias ou deseje relatar problemas, envie uma solicitação via [Issues](https://github.com/seu-usuario/data-cleaning-script/issues) ou abra um pull request.

---

## Autor

- **Gabriel Matuck**
- **E-mail:** [gabriel.matuck1@gmail.com](mailto:gabriel.matuck1@gmail.com)

---

Aproveite e mantenha seus dados organizados e limpos!

