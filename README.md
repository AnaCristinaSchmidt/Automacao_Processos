# Automacao_Processos

Automação de notas fiscais XML (NF-e) para uma planilha de contas a pagar, com painel no navegador.

## Executar no Windows

Com Python 3.9 ou superior instalado, abra o terminal na pasta do projeto:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe painel.py
```

Também é possível abrir `abrir_painel.bat` com dois cliques. O painel usa `http://localhost:8765`.

## Processar pelo terminal

Coloque os XMLs na pasta `entrada`, feche a planilha no Excel e execute:

```powershell
.\.venv\Scripts\python.exe automacao_nf.py
```

Para monitorar a pasta continuamente, acrescente `--vigiar`. Use Ctrl+C para encerrar.

A planilha `contas_a_pagar.xlsx` é criada localmente. Os XMLs processados vão para `processadas`; arquivos rejeitados vão para `erros`. Consulte `processamento.log` para detalhes.

## Dados fictícios para demonstração

```powershell
.\.venv\Scripts\python.exe gerar_notas_exemplo.py
.\.venv\Scripts\python.exe automacao_nf.py
```

Os exemplos são adicionados à mesma planilha usada pelo programa. XMLs, planilhas, logs e ambiente virtual não são versionados.

Veja também [LEIAME.md](LEIAME.md) para as opções de execução no VS Code.
