Attribute VB_Name = "modRegions"
Option Explicit

Public Sub Region_ClearForm()
    Dim ws As Worksheet
    Set ws = GetWs(SH_REGIOES)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Range("B3:B7").ClearContents
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Public Sub Region_SaveFromForm()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim lo As ListObject
    Set ws = GetWs(SH_REGIOES)
    Set lo = ws.ListObjects(TB_REG)

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Unprotect Password:=pwd

    Dim codigo As String
    Dim nome As String
    Dim endereco As String
    Dim supervisor As String
    Dim cap As Variant

    codigo = UCase$(Trim$(CStr(ws.Range("B3").Value)))
    nome = Trim$(CStr(ws.Range("B4").Value))
    endereco = Trim$(CStr(ws.Range("B5").Value))
    supervisor = Trim$(CStr(ws.Range("B6").Value))
    cap = ws.Range("B7").Value

    If Len(codigo) = 0 Then Err.Raise vbObjectError + 200, APP_TITLE, "Codigo da regiao e obrigatorio."
    If Len(nome) = 0 Then Err.Raise vbObjectError + 201, APP_TITLE, "Nome da regiao e obrigatorio."
    If Len(endereco) = 0 Then Err.Raise vbObjectError + 202, APP_TITLE, "Endereco e obrigatorio."
    If Len(supervisor) = 0 Then Err.Raise vbObjectError + 203, APP_TITLE, "Supervisor e obrigatorio."
    If Not IsNumeric(cap) Then Err.Raise vbObjectError + 204, APP_TITLE, "Capacidade maxima invalida."
    If CLng(cap) <= 0 Then Err.Raise vbObjectError + 205, APP_TITLE, "Capacidade maxima deve ser maior que zero."

    Dim rowIdx As Long
    rowIdx = Region_FindRowByCode(lo, codigo)

    Dim lr As ListRow
    If rowIdx = 0 Then
        Set lr = lo.ListRows.Add
    Else
        Set lr = lo.ListRows(rowIdx)
    End If

    With lr.Range
        .Cells(1, TableColIndex(lo, "RegiaoCodigo")).Value = codigo
        .Cells(1, TableColIndex(lo, "RegiaoNome")).Value = nome
        .Cells(1, TableColIndex(lo, "EnderecoCompleto")).Value = endereco
        .Cells(1, TableColIndex(lo, "Supervisor")).Value = supervisor
        .Cells(1, TableColIndex(lo, "CapacidadeMaxima")).Value = CLng(cap)
    End With

    ws.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True

    Setup_RefreshAfterDataChange
    MsgBox "Regiao salva: " & codigo, vbInformation, APP_TITLE
    Exit Sub
ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
    Dim msg As String
    msg = errDesc
    If Len(msg) = 0 Then msg = "Erro " & CStr(errNum)
    MsgBox msg, vbExclamation, APP_TITLE
End Sub

Private Function Region_FindRowByCode(ByVal lo As ListObject, ByVal codigo As String) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim r As Long
    Dim idx As Long
    idx = TableColIndex(lo, "RegiaoCodigo")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idx).Value), codigo, vbTextCompare) = 0 Then
            Region_FindRowByCode = r
            Exit Function
        End If
    Next r
End Function

Public Function Region_GetName(ByVal codigo As String) As String
    Dim lo As ListObject
    Set lo = GetWs(SH_REGIOES).ListObjects(TB_REG)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim idxCode As Long
    Dim idxName As Long
    idxCode = TableColIndex(lo, "RegiaoCodigo")
    idxName = TableColIndex(lo, "RegiaoNome")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxCode).Value), codigo, vbTextCompare) = 0 Then
            Region_GetName = CStr(lo.DataBodyRange.Cells(r, idxName).Value)
            Exit Function
        End If
    Next r
End Function

Public Function Region_GetCapacity(ByVal codigo As String) As Long
    Dim lo As ListObject
    Set lo = GetWs(SH_REGIOES).ListObjects(TB_REG)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim idxCode As Long
    Dim idxCap As Long
    idxCode = TableColIndex(lo, "RegiaoCodigo")
    idxCap = TableColIndex(lo, "CapacidadeMaxima")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxCode).Value), codigo, vbTextCompare) = 0 Then
            Region_GetCapacity = CLng(lo.DataBodyRange.Cells(r, idxCap).Value)
            Exit Function
        End If
    Next r
End Function

