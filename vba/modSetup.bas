Attribute VB_Name = "modSetup"
Option Explicit

Public Sub Setup_InitializeWorkbook()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Setup_CreateConfig
    Setup_CreateDatabaseSheets
    Setup_CreateCadastro
    Setup_CreateRegioes
    Setup_CreateAlocacao
    Setup_CreateConsulta
    Setup_CreateDashboard
    Setup_CreateRelatorio
    Setup_RefreshAfterDataChange
    Setup_ProtectAll

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Estrutura criada/atualizada com sucesso.", vbInformation, APP_TITLE
End Sub

Public Sub Setup_RefreshAfterDataChange()
    On Error GoTo ErrHandler
    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))

    On Error Resume Next
    GetWs(SH_CADASTRO).Unprotect Password:=pwd
    GetWs(SH_ALOC_FORM).Unprotect Password:=pwd
    GetWs(SH_CONSULTA).Unprotect Password:=pwd
    GetWs(SH_DASH).Unprotect Password:=pwd
    On Error GoTo ErrHandler

    Setup_CreateNamedRanges
    Setup_ApplyDataValidation
    Dashboard_RefreshAll
    Setup_ProtectAll
    Exit Sub

ErrHandler:
    On Error Resume Next
    Setup_ProtectAll
End Sub

Private Sub Setup_CreateConfig()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(SH_CONFIG)
    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ClearSheet ws
        ws.Range("A1").Value = "Chave"
        ws.Range("B1").Value = "Valor"
        ws.Range("A2").Value = "SenhaProtecao"
        ws.Range(CFG_PROTECT_PWD_CELL).Value = "alocacao"
        ws.Range("A3").Value = "CodigoAutorizacaoRetroativa"
        ws.Range(CFG_RETRO_CODE_CELL).Value = "RETRO-OK"
        ws.Range("A4").Value = "DiasPermitidosRetroativo"
        ws.Range(CFG_RETRO_ALLOW_DAYS_CELL).Value = 0
        ws.Range("A6").Value = "StatusFuncionarios"
        ws.Range("A7").Value = "Ativo"
        ws.Range("A8").Value = "Inativo"
        ws.Range("D6").Value = "Departamentos"
        ws.Range("D7").Value = "Operacoes"
        ws.Range("D8").Value = "Administrativo"
        ws.Range("D9").Value = "TI"
        ws.Range("D10").Value = "RH"
        ws.Range("G6").Value = "Cargos"
        ws.Range("G7").Value = "Analista"
        ws.Range("G8").Value = "Assistente"
        ws.Range("G9").Value = "Coordenador"
        ws.Range("G10").Value = "Supervisor"
    End If
    ws.Visible = xlSheetVisible
End Sub

Private Sub Setup_CreateDatabaseSheets()
    Dim wsF As Worksheet
    Dim wsA As Worksheet
    Set wsF = EnsureWorksheet(SH_FUNC_DB)
    Set wsA = EnsureWorksheet(SH_ALOC_DB)
    If Not TableExists(wsF, TB_FUNC) And wsF.UsedRange.Count = 1 Then ClearSheet wsF
    If Not TableExists(wsA, TB_ALOC) And wsA.UsedRange.Count = 1 Then ClearSheet wsA

    Dim loF As ListObject
    Set loF = EnsureTable(wsF, TB_FUNC, 1, Array("FuncionarioID", "NomeCompleto", "CPF", "DataAdmissao", "Cargo", "Departamento", "Status", "DataCadastro", "UltimaAtualizacao"))
    loF.ListColumns("FuncionarioID").Range.NumberFormat = "@"
    loF.ListColumns("CPF").Range.NumberFormat = "@"
    loF.ListColumns("DataAdmissao").Range.NumberFormat = "dd/mm/yyyy"
    loF.ListColumns("DataCadastro").Range.NumberFormat = "dd/mm/yyyy hh:mm"
    loF.ListColumns("UltimaAtualizacao").Range.NumberFormat = "dd/mm/yyyy hh:mm"
    wsF.Columns.AutoFit

    Dim loA As ListObject
    Set loA = EnsureTable(wsA, TB_ALOC, 1, Array("AlocacaoID", "FuncionarioID", "RegiaoCodigo", "DataInicio", "DataFim", "Observacoes", "DataRegistro", "Usuario"))
    If Not loA.DataBodyRange Is Nothing Then
        loA.ListColumns("DataInicio").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        loA.ListColumns("DataFim").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    End If
    wsA.Columns.AutoFit
End Sub

Private Sub Setup_CreateCadastro()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(SH_CADASTRO)
    ClearSheet ws

    ws.Columns("A:A").ColumnWidth = 24
    ws.Columns("B:B").ColumnWidth = 42
    ws.Columns("C:C").ColumnWidth = 4
    ws.Columns("D:D").ColumnWidth = 20

    ApplySheetTheme ws, "Cadastro de Funcionarios", "A1:D1"

    ws.Range("A3").Value = "FuncionarioID"
    ws.Range("A4").Value = "NomeCompleto"
    ws.Range("A5").Value = "CPF"
    ws.Range("A6").Value = "DataAdmissao"
    ws.Range("A7").Value = "Cargo"
    ws.Range("A8").Value = "Departamento"
    ws.Range("A9").Value = "Status"

    ws.Range("B3").Value = "(automatico)"
    ws.Range("B3").NumberFormat = "@"
    ws.Range("B5").NumberFormat = "@"
    ws.Range("B6").NumberFormat = "dd/mm/yyyy"

    ws.Cells.Locked = True
    ws.Range("B4:B9").Locked = False

    RemoveShapesByOnAction ws, "Employee_SaveFromForm", "Employee_ClearForm"

    AddSheetButtonAtRange ws, "Salvar/Atualizar", "Employee_SaveFromForm", ws.Range("B11:C12")
    AddSheetButtonAtRange ws, "Limpar", "Employee_ClearForm", ws.Range("D11:D12")

    ws.Range("A3:A9").Font.Bold = True
    ws.Range("A3:A9").VerticalAlignment = xlCenter
    ws.Range("B3:B9").Interior.Color = RGB(242, 242, 242)
    ws.Range("B3:B9").Borders.LineStyle = xlContinuous
    ws.Range("B3:B9").Borders.Color = RGB(220, 220, 220)
End Sub

Private Sub Setup_CreateRegioes()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(SH_REGIOES)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Range("A1:E9").Clear

    ws.Columns("A:A").ColumnWidth = 22
    ws.Columns("B:B").ColumnWidth = 48
    ws.Columns("C:C").ColumnWidth = 4
    ws.Columns("D:D").ColumnWidth = 22
    ws.Columns("E:E").ColumnWidth = 18

    ApplySheetTheme ws, "Regioes", "A1:E1"

    RemoveShapesByOnAction ws, "Region_SaveFromForm", "Region_ClearForm"

    ws.Range("A2").Value = "Cadastro de Regioes"
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 12

    ws.Range("A3").Value = "Codigo"
    ws.Range("A4").Value = "Nome"
    ws.Range("A5").Value = "Endereco"
    ws.Range("A6").Value = "Supervisor"
    ws.Range("A7").Value = "CapacidadeMaxima"

    ws.Range("A3:A7").Font.Bold = True

    ws.Range("B3:B7").Interior.Color = RGB(242, 242, 242)
    ws.Range("B3:B7").Borders.LineStyle = xlContinuous
    ws.Range("B3:B7").Borders.Color = RGB(220, 220, 220)
    ws.Range("B7").NumberFormat = "0"

    ws.Cells.Locked = True
    ws.Range("B3:B7").Locked = False

    AddSheetButtonAtRange ws, "Salvar/Atualizar", "Region_SaveFromForm", ws.Range("B8:C9")
    AddSheetButtonAtRange ws, "Limpar", "Region_ClearForm", ws.Range("D8:D9")

    Dim loR As ListObject
    Set loR = EnsureTable(ws, TB_REG, 10, Array("RegiaoCodigo", "RegiaoNome", "EnderecoCompleto", "Supervisor", "CapacidadeMaxima"))
    ws.Range("A10").EntireRow.RowHeight = 18
    ws.Columns.AutoFit
End Sub

Private Sub Setup_CreateAlocacao()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(SH_ALOC_FORM)
    ClearSheet ws

    ws.Columns("A:A").ColumnWidth = 24
    ws.Columns("B:B").ColumnWidth = 46
    ws.Columns("C:C").ColumnWidth = 4
    ws.Columns("D:D").ColumnWidth = 20

    ApplySheetTheme ws, "Alocacao por Regiao", "A1:D1"

    ws.Range("A3").Value = "Funcionario"
    ws.Range("A4").Value = "Regiao"
    ws.Range("A5").Value = "DataInicio"
    ws.Range("A6").Value = "DataFim"
    ws.Range("A7").Value = "Observacoes"
    ws.Range("A9").Value = "AutorizacaoRetroativa"
    ws.Range("A10").Value = "CodigoAutorizacao"

    ws.Range("B5").NumberFormat = "dd/mm/yyyy"
    ws.Range("B6").NumberFormat = "dd/mm/yyyy"

    ws.Rows("7:7").RowHeight = 60
    ws.Range("B7").WrapText = True

    RemoveShapesByOnAction ws, "Allocation_SaveFromForm", "Allocation_ClearForm"

    ws.Cells.Locked = True
    ws.Range("B3:B7").Locked = False
    ws.Range("B9:B10").Locked = False

    AddSheetButtonAtRange ws, "Salvar Alocacao", "Allocation_SaveFromForm", ws.Range("B12:C13")
    AddSheetButtonAtRange ws, "Limpar", "Allocation_ClearForm", ws.Range("D12:D13")

    ws.Range("A3:A10").Font.Bold = True
    ws.Range("B3:B6").Interior.Color = RGB(242, 242, 242)
    ws.Range("B9:B10").Interior.Color = RGB(242, 242, 242)
    ws.Range("B3:B6").Borders.LineStyle = xlContinuous
    ws.Range("B9:B10").Borders.LineStyle = xlContinuous
    ws.Range("B3:B6").Borders.Color = RGB(220, 220, 220)
    ws.Range("B9:B10").Borders.Color = RGB(220, 220, 220)
End Sub

Private Sub Setup_CreateConsulta()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(SH_CONSULTA)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Cells.Clear

    ws.Columns("A:A").ColumnWidth = 28
    ws.Columns("B:B").ColumnWidth = 30
    ws.Columns("C:C").ColumnWidth = 4
    ws.Columns("D:D").ColumnWidth = 26
    ws.Columns("E:E").ColumnWidth = 14
    ws.Columns("F:F").ColumnWidth = 14

    ApplySheetTheme ws, "Consulta Historica", "A1:F1"

    RemoveShapesByOnAction ws, "Query_Run", "Query_Clear"

    ws.Range("A3").Value = "Funcionario (ID ou Nome)"
    ws.Range("A4").Value = "Regiao (codigo)"
    ws.Range("A5").Value = "DataInicial"
    ws.Range("A6").Value = "DataFinal"

    ws.Range("B5").NumberFormat = "dd/mm/yyyy"
    ws.Range("B6").NumberFormat = "dd/mm/yyyy"

    ws.Cells.Locked = True
    ws.Range("B3:B6").Locked = False

    AddSheetButtonAtRange ws, "Buscar", "Query_Run", ws.Range("B7:C7")
    AddSheetButtonAtRange ws, "Limpar", "Query_Clear", ws.Range("D7:E7")

    Dim loQ As ListObject
    Set loQ = EnsureTable(ws, TB_QUERY, 10, Array("AlocacaoID", "FuncionarioID", "NomeCompleto", "CPF", "RegiaoCodigo", "RegiaoNome", "DataInicio", "DataFim", "Observacoes"))
    ws.Columns.AutoFit

    ws.Range("A3:A6").Font.Bold = True
    ws.Range("B3:B6").Interior.Color = RGB(242, 242, 242)
    ws.Range("B3:B6").Borders.LineStyle = xlContinuous
    ws.Range("B3:B6").Borders.Color = RGB(220, 220, 220)
End Sub

Private Sub Setup_CreateDashboard()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(SH_DASH)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Cells.Clear

    ws.Columns("A:A").ColumnWidth = 30
    ws.Columns("B:B").ColumnWidth = 14
    ws.Columns("C:C").ColumnWidth = 4
    ws.Columns("D:D").ColumnWidth = 28
    ws.Columns("E:E").ColumnWidth = 14
    ws.Columns("F:F").ColumnWidth = 14
    ws.Columns("H:H").ColumnWidth = 16
    ws.Columns("I:I").ColumnWidth = 16
    ws.Columns("J:J").ColumnWidth = 16
    ws.Columns("K:K").ColumnWidth = 16

    ApplySheetTheme ws, "Dashboard", "A1:F1"

    RemoveShapesByOnAction ws, "Dashboard_RefreshAll"

    ws.Range("A3").Value = "Indicadores"
    ws.Range("A3").Font.Bold = True
    ws.Range("A5").Value = "Funcionarios sem alocacao"
    ws.Range("A6").Value = "Alocacoes vencendo (7 dias)"

    ws.Range("B5:B6").Font.Bold = True
    ws.Range("B5:B6").Font.Size = 16
    ws.Range("B5:B6").HorizontalAlignment = xlCenter
    ws.Range("B5:B6").Interior.Color = RGB(242, 242, 242)
    ws.Range("B5:B6").Borders.LineStyle = xlContinuous
    ws.Range("B5:B6").Borders.Color = RGB(220, 220, 220)

    AddSheetButtonAtRange ws, "Atualizar", "Dashboard_RefreshAll", ws.Range("D3:E4")

    Dim loD As ListObject
    Set loD = EnsureTable(ws, TB_DASH, 9, Array("RegiaoCodigo", "RegiaoNome", "CapacidadeMaxima", "AlocadosHoje", "TaxaOcupacao"))
    ws.Columns.AutoFit
End Sub

Private Sub Setup_CreateRelatorio()
    Dim ws As Worksheet
    Set ws = EnsureWorksheet(SH_REL)
    ClearSheet ws
    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub Setup_CreateNamedRanges()
    Dim wsR As Worksheet
    Dim wsF As Worksheet
    Set wsR = GetWs(SH_REGIOES)
    Set wsF = GetWs(SH_FUNC_DB)

    On Error Resume Next
    ThisWorkbook.Names(NAME_REG_CODES).Delete
    ThisWorkbook.Names(NAME_FUNC_IDS).Delete
    On Error GoTo 0

    Dim loR As ListObject
    Dim loF As ListObject
    Set loR = wsR.ListObjects(TB_REG)
    Set loF = wsF.ListObjects(TB_FUNC)

    If Not loR.DataBodyRange Is Nothing Then
        ThisWorkbook.Names.Add Name:=NAME_REG_CODES, RefersTo:=loR.ListColumns("RegiaoCodigo").DataBodyRange
    Else
        ThisWorkbook.Names.Add Name:=NAME_REG_CODES, RefersTo:=wsR.Range("A11:A11")
    End If

    If Not loF.DataBodyRange Is Nothing Then
        ThisWorkbook.Names.Add Name:=NAME_FUNC_IDS, RefersTo:=loF.ListColumns("FuncionarioID").DataBodyRange
    Else
        ThisWorkbook.Names.Add Name:=NAME_FUNC_IDS, RefersTo:=wsF.Range("A2:A2")
    End If
End Sub

Private Sub Setup_ApplyDataValidation()
    Dim wsC As Worksheet
    Dim wsA As Worksheet
    Set wsC = GetWs(SH_CADASTRO)
    Set wsA = GetWs(SH_ALOC_FORM)

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsC.Unprotect Password:=pwd
    wsA.Unprotect Password:=pwd

    wsC.Range("B9").Validation.Delete
    wsC.Range("B9").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=Config!$A$7:$A$8"

    wsC.Range("B8").Validation.Delete
    wsC.Range("B8").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=Config!$D$7:$D$10"

    wsC.Range("B7").Validation.Delete
    wsC.Range("B7").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=Config!$G$7:$G$10"

    wsA.Range("B3").Validation.Delete
    wsA.Range("B3").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=" & NAME_FUNC_IDS

    wsA.Range("B4").Validation.Delete
    wsA.Range("B4").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=" & NAME_REG_CODES

    wsA.Range("B9").Validation.Delete
    wsA.Range("B9").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "SIM,NAO"
    wsA.Range("B9").Value = "NAO"

    wsC.Protect Password:=pwd, UserInterfaceOnly:=True
    wsA.Protect Password:=pwd, UserInterfaceOnly:=True
End Sub

Private Sub Setup_ProtectAll()
    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=pwd
    Next ws

    GetWs(SH_FUNC_DB).Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    GetWs(SH_ALOC_DB).Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    GetWs(SH_REGIOES).Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    GetWs(SH_CONSULTA).Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    GetWs(SH_DASH).Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    GetWs(SH_CADASTRO).Protect Password:=pwd, UserInterfaceOnly:=True
    GetWs(SH_ALOC_FORM).Protect Password:=pwd, UserInterfaceOnly:=True
End Sub

