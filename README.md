# HomeZy - Automação de Extração e Tratamento de Dados

Pipeline de automação desenvolvido em **Python** para executar **web scraping autenticado**, baixar relatórios de sistemas web, tratar arquivos Excel e atualizar bases analíticas de forma automática.

## O que o projeto faz

- acessa sistemas web com login
- navega até relatórios específicos
- baixa arquivos automaticamente
- converte `.xls` para `.xlsx`
- trata e padroniza dados
- atualiza bases finais em Excel
- executa tudo em sequência por um pipeline único
- roda diariamente via **Agendador de Tarefas do Windows**

## Principais tecnologias utilizadas

- **Python**
- **Selenium**
- **Pandas**
- **OpenPyXL**
- **PyWin32**
- **python-dotenv**
- **Windows Task Scheduler**
- **Batch script (.bat)**

## Destaques técnicos

- automação de múltiplos sistemas web com autenticação
- detecção robusta de downloads e renomeação padronizada
- manipulação de planilhas Excel com preservação de estrutura
- substituição de fórmulas do Excel por regras de negócio em Python
- execução automática com logs por data e hora

## Estrutura principal

- `extrator_dados.py` -> extração de relatórios Olist
- `extrator_dados3.py` -> extração de relatório Uoou
- `tratamento_dados.py` -> atualização da base `fEcommerce.xlsx`
- `tratamento_dados2.py` -> atualização da base `resumo_proforma_pedido.xlsx`
- `tratamento_dados3.py` -> atualização da base `Arquivo - Data de entrega pedidos.xlsx`
- `main.py` -> orquestração do pipeline
- `rodar_homezy.bat` -> execução automatizada com log

## Resultado

A automação reduz trabalho manual, padroniza a atualização das bases e aumenta a confiabilidade do processo operacional.