Attribute VB_Name = "modEmployees"
Option Explicit

Public Sub Employee_ClearForm()
    Dim ws As Worksheet
    Set ws = GetWs(SH_CADASTRO)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Range("B3:B9").ClearContents
    ws.Range("B3").Value = "(automatico)"
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True
End Sub

Public Sub Employee_SaveFromForm()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim lo As ListObject
    Set ws = GetWs(SH_CADASTRO)
    Dim wsDb As Worksheet
    Set wsDb = GetWs(SH_FUNC_DB)
    Set lo = wsDb.ListObjects(TB_FUNC)

    Dim empId As String
    Dim nome As String
    Dim cpf As String
    Dim dtAdm As Variant
    Dim cargo As String
    Dim dept As String
    Dim status As String

    empId = Trim$(CStr(ws.Range("B3").Value))
    nome = Trim$(CStr(ws.Range("B4").Value))
    cpf = Trim$(CStr(ws.Range("B5").Value))
    dtAdm = ws.Range("B6").Value
    cargo = Trim$(CStr(ws.Range("B7").Value))
    dept = Trim$(CStr(ws.Range("B8").Value))
    status = Trim$(CStr(ws.Range("B9").Value))

    If Len(nome) = 0 Then Err.Raise vbObjectError + 100, APP_TITLE, "Nome completo e obrigatorio."
    If Len(cpf) = 0 Then Err.Raise vbObjectError + 101, APP_TITLE, "CPF e obrigatorio."
    If Not IsValidCPF(cpf) Then Err.Raise vbObjectError + 102, APP_TITLE, "CPF invalido."
    If Not IsDate(dtAdm) Then Err.Raise vbObjectError + 103, APP_TITLE, "Data de admissao invalida."
    If Len(cargo) = 0 Then Err.Raise vbObjectError + 104, APP_TITLE, "Cargo e obrigatorio."
    If Len(dept) = 0 Then Err.Raise vbObjectError + 105, APP_TITLE, "Departamento e obrigatorio."
    If Len(status) = 0 Then Err.Raise vbObjectError + 106, APP_TITLE, "Status e obrigatorio."

    Dim existingRow As ListRow
    Dim rowIdx As Long

    rowIdx = Employee_FindRowById(lo, empId)
    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsDb.Unprotect Password:=pwd

    If rowIdx = 0 Then
        If empId = "(automatico)" Or Len(empId) = 0 Then empId = Employee_NextId(lo)
        If Employee_FindRowByCPF(lo, cpf) <> 0 Then Err.Raise vbObjectError + 107, APP_TITLE, "Ja existe funcionario com este CPF."
        Set existingRow = lo.ListRows.Add
    Else
        Set existingRow = lo.ListRows(rowIdx)
        Dim otherCpfRow As Long
        otherCpfRow = Employee_FindRowByCPF(lo, cpf)
        If otherCpfRow <> 0 And otherCpfRow <> rowIdx Then Err.Raise vbObjectError + 108, APP_TITLE, "CPF ja pertence a outro funcionario."
    End If

    Dim digits As String
    digits = NormalizeDigits(cpf)

    With existingRow.Range
        .Cells(1, TableColIndex(lo, "FuncionarioID")).Value = empId
        .Cells(1, TableColIndex(lo, "NomeCompleto")).Value = nome
        .Cells(1, TableColIndex(lo, "CPF")).NumberFormat = "@"
        .Cells(1, TableColIndex(lo, "CPF")).Value = "'" & digits
        .Cells(1, TableColIndex(lo, "DataAdmissao")).Value = CDate(dtAdm)
        .Cells(1, TableColIndex(lo, "Cargo")).Value = cargo
        .Cells(1, TableColIndex(lo, "Departamento")).Value = dept
        .Cells(1, TableColIndex(lo, "Status")).Value = status
        If rowIdx = 0 Then .Cells(1, TableColIndex(lo, "DataCadastro")).Value = Now
        .Cells(1, TableColIndex(lo, "UltimaAtualizacao")).Value = Now
    End With

    wsDb.Columns.AutoFit

    wsDb.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True

    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Range("B3").Value = empId
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True

    Setup_RefreshAfterDataChange
    MsgBox "Funcionario salvo: " & empId, vbInformation, APP_TITLE
    Exit Sub
ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not wsDb Is Nothing Then
        Dim p As String
        p = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
        wsDb.Protect Password:=p, UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    Dim msg As String
    msg = errDesc
    If Len(msg) = 0 Then msg = "Erro " & CStr(errNum)
    MsgBox msg, vbExclamation, APP_TITLE
End Sub

Private Function Employee_FindRowById(ByVal lo As ListObject, ByVal empId As String) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Len(Trim$(empId)) = 0 Then Exit Function

    Dim r As Long
    Dim idxId As Long
    idxId = TableColIndex(lo, "FuncionarioID")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxId).Value), empId, vbTextCompare) = 0 Then
            Employee_FindRowById = r
            Exit Function
        End If
    Next r
End Function

Private Function Employee_FindRowByCPF(ByVal lo As ListObject, ByVal cpf As String) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim d As String
    d = NormalizeDigits(cpf)

    Dim r As Long
    Dim idxCpf As Long
    idxCpf = TableColIndex(lo, "CPF")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If CStr(lo.DataBodyRange.Cells(r, idxCpf).Value) = d Then
            Employee_FindRowByCPF = r
            Exit Function
        End If
    Next r
End Function

Private Function Employee_NextId(ByVal lo As ListObject) As String
    Dim maxN As Long
    Dim idxId As Long
    idxId = TableColIndex(lo, "FuncionarioID")
    maxN = 0

    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        Dim v As String
        Dim n As Long
        For r = 1 To lo.DataBodyRange.Rows.Count
            v = CStr(lo.DataBodyRange.Cells(r, idxId).Value)
            If Len(v) >= 2 Then
                n = Val(Mid$(v, 2))
                If n > maxN Then maxN = n
            End If
        Next r
    End If

    Employee_NextId = "F" & Format$(maxN + 1, "000000")
End Function

Public Function Employee_IsActive(ByVal funcionarioId As String) As Boolean
    Dim lo As ListObject
    Set lo = GetWs(SH_FUNC_DB).ListObjects(TB_FUNC)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim idxId As Long
    Dim idxStatus As Long
    idxId = TableColIndex(lo, "FuncionarioID")
    idxStatus = TableColIndex(lo, "Status")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxId).Value), funcionarioId, vbTextCompare) = 0 Then
            Employee_IsActive = (StrComp(CStr(lo.DataBodyRange.Cells(r, idxStatus).Value), "Ativo", vbTextCompare) = 0)
            Exit Function
        End If
    Next r
End Function

Public Function Employee_GetName(ByVal funcionarioId As String) As String
    Dim lo As ListObject
    Set lo = GetWs(SH_FUNC_DB).ListObjects(TB_FUNC)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim idxId As Long
    Dim idxNome As Long
    idxId = TableColIndex(lo, "FuncionarioID")
    idxNome = TableColIndex(lo, "NomeCompleto")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxId).Value), funcionarioId, vbTextCompare) = 0 Then
            Employee_GetName = CStr(lo.DataBodyRange.Cells(r, idxNome).Value)
            Exit Function
        End If
    Next r
End Function

