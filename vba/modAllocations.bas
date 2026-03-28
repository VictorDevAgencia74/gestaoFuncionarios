Attribute VB_Name = "modAllocations"
Option Explicit

Public Sub Allocation_ClearForm()
    Dim ws As Worksheet
    Set ws = GetWs(SH_ALOC_FORM)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Range("B3:B7").ClearContents
    ws.Range("B9").Value = "NAO"
    ws.Range("B10").ClearContents
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True
End Sub

Public Sub Allocation_SaveFromForm()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = GetWs(SH_ALOC_FORM)

    Dim funcionarioId As String
    Dim regiaoCodigo As String
    Dim dtIniV As Variant
    Dim dtFimV As Variant
    Dim obs As String
    Dim authFlag As String
    Dim authCode As String

    funcionarioId = Trim$(CStr(ws.Range("B3").Value))
    regiaoCodigo = UCase$(Trim$(CStr(ws.Range("B4").Value)))
    dtIniV = ws.Range("B5").Value
    dtFimV = ws.Range("B6").Value
    obs = Trim$(CStr(ws.Range("B7").Value))
    authFlag = UCase$(Trim$(CStr(ws.Range("B9").Value)))
    authCode = Trim$(CStr(ws.Range("B10").Value))

    If Len(funcionarioId) = 0 Then Err.Raise vbObjectError + 300, APP_TITLE, "Funcionario e obrigatorio."
    If Len(regiaoCodigo) = 0 Then Err.Raise vbObjectError + 301, APP_TITLE, "Regiao e obrigatoria."
    If Not IsDate(dtIniV) Then Err.Raise vbObjectError + 302, APP_TITLE, "Data de inicio invalida."
    If Not IsDate(dtFimV) Then Err.Raise vbObjectError + 303, APP_TITLE, "Data de termino invalida."

    Dim dtIni As Date
    Dim dtFim As Date
    dtIni = CDate(dtIniV)
    dtFim = CDate(dtFimV)
    If dtIni > dtFim Then Err.Raise vbObjectError + 304, APP_TITLE, "Data de inicio nao pode ser maior que data de termino."

    If Not Employee_IsActive(funcionarioId) Then Err.Raise vbObjectError + 305, APP_TITLE, "Funcionario inativo ou inexistente."
    If Region_GetCapacity(regiaoCodigo) <= 0 Then Err.Raise vbObjectError + 306, APP_TITLE, "Regiao inexistente ou sem capacidade configurada."

    Allocation_ValidateRetroactive dtIni, authFlag, authCode
    Allocation_ValidateNoOverlap funcionarioId, dtIni, dtFim
    Allocation_ValidateCapacity regiaoCodigo, dtIni, dtFim

    Dim lo As ListObject
    Dim wsDb As Worksheet
    Set wsDb = GetWs(SH_ALOC_DB)
    Set lo = wsDb.ListObjects(TB_ALOC)

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsDb.Unprotect Password:=pwd

    Dim lr As ListRow
    Set lr = lo.ListRows.Add
    With lr.Range
        .Cells(1, TableColIndex(lo, "AlocacaoID")).Value = "A-" & NewGuidId()
        .Cells(1, TableColIndex(lo, "FuncionarioID")).Value = funcionarioId
        .Cells(1, TableColIndex(lo, "RegiaoCodigo")).Value = regiaoCodigo
        .Cells(1, TableColIndex(lo, "DataInicio")).Value = dtIni
        .Cells(1, TableColIndex(lo, "DataFim")).Value = dtFim
        .Cells(1, TableColIndex(lo, "Observacoes")).Value = obs
        .Cells(1, TableColIndex(lo, "DataRegistro")).Value = Now
        .Cells(1, TableColIndex(lo, "Usuario")).Value = Application.UserName
    End With

    wsDb.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True

    Dashboard_RefreshAll
    MsgBox "Alocacao salva para " & funcionarioId & " em " & regiaoCodigo, vbInformation, APP_TITLE
    Exit Sub
ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not wsDb Is Nothing Then
        wsDb.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    Dim msg As String
    msg = errDesc
    If Len(msg) = 0 Then msg = "Erro " & CStr(errNum)
    MsgBox msg, vbExclamation, APP_TITLE
End Sub

Private Sub Allocation_ValidateRetroactive(ByVal dtIni As Date, ByVal authFlag As String, ByVal authCode As String)
    Dim allowed As Long
    allowed = CLng(GetConfigValue(CFG_RETRO_ALLOW_DAYS_CELL))

    If dtIni >= Date - allowed Then Exit Sub

    If authFlag <> "SIM" Then Err.Raise vbObjectError + 320, APP_TITLE, "Alocacao retroativa requer autorizacao."
    If StrComp(authCode, CStr(GetConfigValue(CFG_RETRO_CODE_CELL)), vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 321, APP_TITLE, "Codigo de autorizacao invalido."
    End If
End Sub

Private Sub Allocation_ValidateNoOverlap(ByVal funcionarioId As String, ByVal dtIni As Date, ByVal dtFim As Date)
    Dim lo As ListObject
    Set lo = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long
    Dim idxEmp As Long
    Dim idxIni As Long
    Dim idxFim As Long
    idxEmp = TableColIndex(lo, "FuncionarioID")
    idxIni = TableColIndex(lo, "DataInicio")
    idxFim = TableColIndex(lo, "DataFim")

    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxEmp).Value), funcionarioId, vbTextCompare) = 0 Then
            If DateRangesOverlap(dtIni, dtFim, CDate(lo.DataBodyRange.Cells(r, idxIni).Value), CDate(lo.DataBodyRange.Cells(r, idxFim).Value)) Then
                Err.Raise vbObjectError + 330, APP_TITLE, "Sobreposicao de periodos para o mesmo funcionario."
            End If
        End If
    Next r
End Sub

Private Sub Allocation_ValidateCapacity(ByVal regiaoCodigo As String, ByVal dtIni As Date, ByVal dtFim As Date)
    Dim cap As Long
    cap = Region_GetCapacity(regiaoCodigo)
    If cap <= 0 Then Exit Sub

    Dim lo As ListObject
    Set lo = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)
    Dim countOverlap As Long
    countOverlap = 0

    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        Dim idxReg As Long
        Dim idxIni As Long
        Dim idxFim As Long
        idxReg = TableColIndex(lo, "RegiaoCodigo")
        idxIni = TableColIndex(lo, "DataInicio")
        idxFim = TableColIndex(lo, "DataFim")
        For r = 1 To lo.DataBodyRange.Rows.Count
            If StrComp(CStr(lo.DataBodyRange.Cells(r, idxReg).Value), regiaoCodigo, vbTextCompare) = 0 Then
                If DateRangesOverlap(dtIni, dtFim, CDate(lo.DataBodyRange.Cells(r, idxIni).Value), CDate(lo.DataBodyRange.Cells(r, idxFim).Value)) Then
                    countOverlap = countOverlap + 1
                End If
            End If
        Next r
    End If

    If countOverlap + 1 > cap Then
        Err.Raise vbObjectError + 340, APP_TITLE, "Capacidade maxima excedida para o periodo informado."
    End If
End Sub

