## Gestão de Alocação de Funcionários por Região (Excel + VBA)

Este repositório contém módulos VBA que montam (do zero) uma solução em Excel para:

- Cadastro de funcionários com ID único e validações
- Cadastro de regiões com capacidade máxima
- Alocação por período, sem sobreposição e com checagem de capacidade
- Consulta histórica com filtros
- Dashboard com indicadores e gráficos
- Relatórios mensais em PDF

### Como usar

1. Abra o arquivo Excel onde você quer usar a solução (recomendado: salvar como `.xlsm`).
2. Abra o Editor do VBA (ALT+F11).
3. Importe os módulos da pasta `vba/`:
   - Menu `File` → `Import File...` → selecione todos os `.bas`.
4. Volte ao Excel e execute a macro `Setup_InitializeWorkbook`.

### Macros principais

- `Setup_InitializeWorkbook`: cria planilhas, tabelas, validações, botões e proteção.
- `Sample_GenerateData`: cria dados simulados (>= 50 funcionários e 10 regiões) e alocações.
- `Employee_SaveFromForm`: grava/atualiza funcionário a partir da aba `Cadastro`.
- `Region_SaveFromForm`: grava/atualiza região a partir da aba `Regiões`.
- `Allocation_SaveFromForm`: grava alocação a partir da aba `Alocação`.
- `Query_Run`: executa consulta histórica na aba `Consulta`.
- `Dashboard_RefreshAll`: atualiza pivôs/indicadores/gráficos.
- `Reports_GenerateMonthlyPDFs`: gera PDFs mensais em `reports/`.
- `Backup_CreateNow`: cria um backup do arquivo na pasta `bkp/`.
- `Backup_Import`: importa um backup (substitui dados e config).

### Backup automático ao fechar

Para perguntar se deseja fazer backup toda vez que fechar o arquivo, adicione este evento no módulo `ThisWorkbook` (VBA Editor → Microsoft Excel Objects → ThisWorkbook):

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Backup_HandleBeforeClose Cancel
End Sub
```

### Observações

- Para exportar PDF, o Excel precisa ter permissão de escrita na pasta do arquivo.
- Proteções são aplicadas automaticamente; a senha fica em `Config!B2` (aba oculta).
