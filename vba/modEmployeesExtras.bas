Attribute VB_Name = "modEmployeesExtras"
Option Explicit

Public Function Employee_GetCPF(ByVal funcionarioId As String) As String
    Dim lo As ListObject
    Set lo = GetWs(SH_FUNC_DB).ListObjects(TB_FUNC)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim idxId As Long
    Dim idxCpf As Long
    idxId = TableColIndex(lo, "FuncionarioID")
    idxCpf = TableColIndex(lo, "CPF")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxId).Value), funcionarioId, vbTextCompare) = 0 Then
            Employee_GetCPF = CStr(lo.DataBodyRange.Cells(r, idxCpf).Value)
            Exit Function
        End If
    Next r
End Function

Public Sub Employee_NormalizeCpfColumn()
    On Error GoTo ErrHandler
    Dim wsDb As Worksheet
    Set wsDb = GetWs(SH_FUNC_DB)
    Dim lo As ListObject
    Set lo = wsDb.ListObjects(TB_FUNC)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsDb.Unprotect Password:=pwd

    Dim idxCpf As Long
    idxCpf = TableColIndex(lo, "CPF")

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim v As String
        v = CStr(lo.DataBodyRange.Cells(r, idxCpf).Value)
        v = NormalizeDigits(v)
        lo.DataBodyRange.Cells(r, idxCpf).NumberFormat = "@"
        lo.DataBodyRange.Cells(r, idxCpf).Value = "'" & v
    Next r

    wsDb.Columns.AutoFit
    wsDb.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    MsgBox "CPF normalizado com sucesso.", vbInformation, APP_TITLE
    Exit Sub

ErrHandler:
    On Error Resume Next
    wsDb.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
    MsgBox "Erro ao normalizar CPF: " & Err.Description, vbExclamation, APP_TITLE
End Sub

