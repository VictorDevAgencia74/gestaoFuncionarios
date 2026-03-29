Attribute VB_Name = "modTests"
Option Explicit

Public Sub Test_RunAll()
    On Error GoTo ErrHandler

    Setup_InitializeWorkbook
    Sample_GenerateData

    Test_EmployeeCannotOverlap
    Test_CapacityBySimultaneousOccupancy
    Test_DashboardCountsOnlyActiveExpiries

    MsgBox "Testes concluidos com sucesso.", vbInformation, APP_TITLE
    Exit Sub

ErrHandler:
    MsgBox "Teste falhou: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Test_EmployeeCannotOverlap()
    Dim wsA As Worksheet
    Dim beforeCount As Long
    Set wsA = GetWs(SH_ALOC_FORM)

    Test_FillAllocationForm wsA, "", "F000001", "R01", Date, Date + 10, "Teste 1", "NAO", ""
    Allocation_SaveFromForm

    beforeCount = Test_TableRowCount(GetWs(SH_ALOC_DB).ListObjects(TB_ALOC))
    Test_FillAllocationForm wsA, "", "F000001", "R02", Date + 5, Date + 12, "Teste sobreposicao", "NAO", ""
    Allocation_SaveFromForm
    Test_AssertTrue Test_TableRowCount(GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)) = beforeCount, "Alocacao sobreposta foi gravada indevidamente."
End Sub

Private Sub Test_CapacityBySimultaneousOccupancy()
    Dim wsR As Worksheet
    Dim loR As ListObject
    Dim rowIdx As Long
    Set wsR = GetWs(SH_REGIOES)
    Set loR = wsR.ListObjects(TB_REG)

    rowIdx = Test_FindTableRow(loR, "RegiaoCodigo", "R10")
    Test_AssertTrue rowIdx > 0, "Regiao R10 nao encontrada para teste."
    loR.DataBodyRange.Cells(rowIdx, TableColIndex(loR, "CapacidadeMaxima")).Value = 1

    Dim wsA As Worksheet
    Dim beforeCount As Long
    Set wsA = GetWs(SH_ALOC_FORM)
    Test_FillAllocationForm wsA, "", "F000002", "R10", Date, Date + 2, "Capacidade 1", "NAO", ""
    Allocation_SaveFromForm

    Test_FillAllocationForm wsA, "", "F000003", "R10", Date + 3, Date + 5, "Sem sobreposicao", "NAO", ""
    Allocation_SaveFromForm

    beforeCount = Test_TableRowCount(GetWs(SH_ALOC_DB).ListObjects(TB_ALOC))
    Test_FillAllocationForm wsA, "", "F000004", "R10", Date + 1, Date + 1, "Deve exceder", "NAO", ""
    Allocation_SaveFromForm
    Test_AssertTrue Test_TableRowCount(GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)) = beforeCount, "Capacidade excedida nao bloqueou a gravacao."
End Sub

Private Sub Test_DashboardCountsOnlyActiveExpiries()
    Dim wsDb As Worksheet
    Dim loA As ListObject
    Dim lr As ListRow
    Dim beforeValue As Long
    Set wsDb = GetWs(SH_ALOC_DB)
    Set loA = wsDb.ListObjects(TB_ALOC)
    SetConfigValue CFG_EXPIRY_WARN_DAYS_CELL, 5
    Dashboard_RefreshAll
    beforeValue = CLng(GetWs(SH_DASH).Range("B6").Value)

    Set lr = loA.ListRows.Add
    With lr.Range
        .Cells(1, TableColIndex(loA, "AlocacaoID")).Value = "A-TEST-FUTURE"
        .Cells(1, TableColIndex(loA, "FuncionarioID")).Value = "F000005"
        .Cells(1, TableColIndex(loA, "RegiaoCodigo")).Value = "R03"
        .Cells(1, TableColIndex(loA, "DataInicio")).Value = Date + 2
        .Cells(1, TableColIndex(loA, "DataFim")).Value = Date + 3
        .Cells(1, TableColIndex(loA, "Observacoes")).Value = "Futura"
        .Cells(1, TableColIndex(loA, "DataRegistro")).Value = Now
        .Cells(1, TableColIndex(loA, "Usuario")).Value = Application.UserName
    End With

    Dashboard_RefreshAll
    Test_AssertTrue CLng(GetWs(SH_DASH).Range("B6").Value) = beforeValue, "Dashboard esta contando alocacoes futuras como vencendo."
    loA.ListRows(loA.ListRows.Count).Delete
    Dashboard_RefreshAll
End Sub
Private Sub Test_FillAllocationForm(ByVal wsA As Worksheet, ByVal alocId As String, ByVal funcId As String, ByVal regCode As String, ByVal dtIni As Date, ByVal dtFim As Date, ByVal obs As String, ByVal authFlag As String, ByVal authCode As String)
    wsA.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsA.Range("B2").Value = alocId
    wsA.Range("B3").Value = funcId
    wsA.Range("B4").Value = regCode
    wsA.Range("B5").Value = dtIni
    wsA.Range("B6").Value = dtFim
    wsA.Range("B7").Value = obs
    wsA.Range("B9").Value = authFlag
    wsA.Range("B10").Value = authCode
    wsA.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True
End Sub

Private Sub Test_AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 980, APP_TITLE, message
End Sub

Private Function Test_FindTableRow(ByVal lo As ListObject, ByVal colName As String, ByVal lookupValue As String) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim idx As Long
    Dim r As Long
    idx = TableColIndex(lo, colName)
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idx).Value), lookupValue, vbTextCompare) = 0 Then
            Test_FindTableRow = r
            Exit Function
        End If
    Next r
End Function

Private Function Test_TableRowCount(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Test_TableRowCount = lo.DataBodyRange.Rows.Count
End Function

