# Automação: nota fiscal → contas a pagar

Coloque o XML da nota fiscal na pasta `entrada`. A linha aparece sozinha na planilha `contas_a_pagar.xlsx`, já com fornecedor, CNPJ, número da nota, parcela, vencimento e valor.

## Primeira vez no VS Code (uma vez só)

1. **Arquivo → Abrir Pasta…** e escolha `Automacao_Processos`. Se o VS Code perguntar se confia nos autores, clique em **Sim**.
2. **Terminal → Executar Tarefa…** e escolha **Preparar ambiente (criar .venv e instalar openpyxl)**. No fim deve aparecer `OK: Python 3.x | openpyxl 3.x`.
3. Se o VS Code perguntar qual interpretador usar, escolha o que tem **.venv** no caminho. Ou use `Ctrl+Shift+P` → **Python: Select Interpreter** → `.venv`.

## Rodar

Abra a aba **Executar e Depurar** (`Ctrl+Shift+D`), escolha a opção na lista do topo e aperte **F5**:

| Opção | O que faz |
|---|---|
| 1. Gerar notas de exemplo | Cria 8 arquivos fictícios na pasta `entrada`, para testar |
| 2. Processar notas da pasta entrada | Lança o que estiver na pasta e termina |
| 3. Vigiar pasta (fica rodando) | Lança cada nota que chegar, a cada 5 segundos. Pare com o botão quadrado vermelho |

Para testar do zero: **Terminal → Executar Tarefa… → Limpar teste**, depois as opções 1 e 2.

## O que acontece com cada arquivo

| Situação | Resultado |
|---|---|
| XML de NF-e válido | Uma linha por parcela na planilha; o arquivo vai para `processadas` |
| Nota já lançada (mesma chave de acesso) | Ignorada, sem duplicar |
| Nota sem parcelas | Lançada com situação **Verificar** |
| PDF, imagem ou XML que não é NF-e | Vai para `erros`, com o motivo em `processamento.log` |

## A planilha

- **Contas a pagar**: a coluna *Situação* se calcula sozinha, com cores: Vencida, Vence em 7 dias, A vencer, Verificar e Paga. Para baixar uma parcela, marque **Sim** na coluna *Pago?*.
- **Resumo**: totais por situação e o total em aberto.
- Feche a planilha no Excel antes de rodar. Se ela estiver aberta, o script avisa e não salva.

## Problemas comuns

- **`py` não é reconhecido**: o Python não está no PATH. Reinstale pelo python.org marcando **Add python.exe to PATH**, ou troque `py -3` por `python` na tarefa (arquivo `.vscode/tasks.json`).
- **Execução de scripts desabilitada no PowerShell** (ao ativar o `.venv`): rode uma vez no terminal `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` e confirme com `S`.
- **`ModuleNotFoundError: openpyxl`**: o VS Code está usando outro Python. Selecione o interpretador `.venv` (passo 3).

## Deixar automático no cliente

- **Agendador de Tarefas do Windows**: rode `C:\caminho\Automacao_Processos\.venv\Scripts\pythonw.exe C:\caminho\Automacao_Processos\automacao_nf.py` a cada 10 minutos.
- **Pasta compartilhada**: mude `PASTA_ENTRADA` no topo de `automacao_nf.py` para uma pasta do OneDrive ou da rede.
- **Direto do e-mail**: uma regra do Outlook que salve anexos `.xml` nessa pasta fecha o ciclo.
