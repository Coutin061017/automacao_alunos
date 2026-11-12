# Automação de Cadastro de Alunos

Projeto de RPA desenvolvido em Python para validar dados de alunos, automatizar inscrições em uma página HTML local e gerar relatórios de acompanhamento. O processo combina automação web, tratamento de dados e indicadores para análise no Looker Studio.

## Como funciona

1. **Leitura:** o Pandas importa a planilha `alunos_projeto_final.xlsx` com as colunas Nome, CPF e WhatsApp.
2. **Validação:** registros com campos obrigatórios vazios ou CPF duplicado no lote são classificados como pendências e não seguem para cadastro.
3. **Cadastro:** o Selenium abre o formulário local, preenche os dados dos alunos válidos e aciona o botão Cadastrar.
4. **Conferência:** o cadastro só é considerado concluído quando aparece na tabela com os dados esperados. O BeautifulSoup extrai o histórico final para confirmar os resultados.
5. **Relatórios:** o sistema separa sucessos e pendências, calcula as taxas de processamento e salva arquivos Excel.

As operações e falhas são registradas em `logs/automacao.log`. Uma cópia da planilha de entrada é preservada em `backup/` com `shutil.copy2()`.

## Tecnologias

Python 3, Pandas, openpyxl, Selenium WebDriver, Google Chrome, BeautifulSoup4, logging, pathlib, os e shutil. Os indicadores podem ser apresentados no Google Looker Studio a partir de dados disponibilizados no Google Sheets.

## Estrutura do projeto

```text
projeto_automacao_alunos/
├── dados/
│   └── alunos_projeto_final.xlsx
├── web/
│   └── index.html
├── drivers/
│   └── chromedriver.exe
├── backup/
├── logs/
├── relatorios/
├── main.py
├── requirements.txt
└── .gitignore
```

As pastas `backup/`, `logs/` e `relatorios/` são utilizadas durante a execução. O driver e a planilha devem estar disponíveis nos caminhos indicados.

## Instalação e execução

No PowerShell, a partir da pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Instale um ChromeDriver compatível com a versão do Google Chrome em `drivers/chromedriver.exe`.

```powershell
python main.py --so-validar  # Apenas valida a planilha
python main.py               # Executa a automação completa
```

## Arquivos gerados

Cada execução completa cria uma subpasta em `relatorios/`, identificada por data e horário.

| Arquivo | Conteúdo |
|---|---|
| `relatorio_sucesso_historico.xlsx` | Cadastros confirmados no histórico HTML. |
| `relatorio_pendencias.xlsx` | Alunos não cadastrados e motivos da pendência. |
| `base_consolidada.xlsx` | Resultado por aluno e indicadores gerais. |
| `base_painel.xlsx` | Dados resumidos para análise no dashboard. |

A taxa de sucesso corresponde a **cadastros confirmados / total de alunos × 100**. A taxa de pendência utiliza a mesma fórmula com o número de pendências.

## Dashboard

Os resultados podem ser acompanhados no painel desenvolvido no Looker Studio:

**[Visualizar dashboard de cadastro de alunos](https://datastudio.google.com/reporting/f3046058-13f5-42c9-b507-064666a5613a)**

A atualização do Google Sheets e do Looker Studio não é automática nesta versão do script: a base gerada precisa ser importada ou sincronizada separadamente.

## Segurança e observações

Utilize dados fictícios na demonstração. Não publique no GitHub planilhas com CPF ou WhatsApp reais, backups, relatórios, logs ou credenciais de acesso. O histórico do formulário existe apenas durante a sessão do navegador; os resultados persistentes são os arquivos gerados pelo programa.

Projeto desenvolvido para fins educacionais, com foco em validação, automação de processos, registro de erros e análise de resultados.
