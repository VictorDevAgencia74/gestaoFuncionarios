Attribute VB_Name = "modAllocations"
Option Explicit

Public Sub Allocation_ClearForm()
    Dim ws As Worksheet
    Set ws = GetWs(SH_ALOC_FORM)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Range("B2").ClearContents
    ws.Range("B3:D7").ClearContents
    ws.Range("B9:D10").ClearContents
    ws.Range("B9").Value = "NAO"
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True
End Sub

Public Sub Allocation_SaveFromForm()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = GetWs(SH_ALOC_FORM)

    Dim alocId As String
    alocId = Trim$(CStr(ws.Range("B2").Value))

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

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))

    If Not Employee_IsActive(funcionarioId) Then Err.Raise vbObjectError + 305, APP_TITLE, "Funcionario inativo ou inexistente."
    If Region_GetCapacity(regiaoCodigo) <= 0 Then Err.Raise vbObjectError + 306, APP_TITLE, "Regiao inexistente ou sem capacidade configurada."

    Dim lo As ListObject
    Dim wsDb As Worksheet
    Set wsDb = GetWs(SH_ALOC_DB)
    Set lo = wsDb.ListObjects(TB_ALOC)

    Dim autoClosed As Boolean
    Dim autoRow As Long
    Dim autoOldFim As Date
    Dim autoNewFim As Date
    Dim autoCandidateFound As Boolean

    Allocation_ValidateRetroactive dtIni, authFlag, authCode
    wsDb.Unprotect Password:=pwd

    Dim rowIdx As Long
    rowIdx = 0
    If Len(alocId) > 0 Then rowIdx = Allocation_FindRowById(lo, alocId)

    If rowIdx = 0 Then
        autoCandidateFound = Allocation_FindAutoCloseCandidate(lo, funcionarioId, dtIni, dtFim, autoRow, autoOldFim, autoNewFim)
        Allocation_ValidateNoOverlapWithCandidate lo, funcionarioId, dtIni, dtFim, autoRow, autoCandidateFound, autoNewFim
        Allocation_ValidateCapacity regiaoCodigo, dtIni, dtFim

        If autoCandidateFound Then
            lo.DataBodyRange.Cells(autoRow, TableColIndex(lo, "DataFim")).Value = autoNewFim
            autoClosed = True
        End If

        Dim lr As ListRow
        Set lr = lo.ListRows.Add
        alocId = "A-" & NewGuidId()
        With lr.Range
            .Cells(1, TableColIndex(lo, "AlocacaoID")).Value = alocId
            .Cells(1, TableColIndex(lo, "FuncionarioID")).Value = funcionarioId
            .Cells(1, TableColIndex(lo, "RegiaoCodigo")).Value = regiaoCodigo
            .Cells(1, TableColIndex(lo, "DataInicio")).Value = dtIni
            .Cells(1, TableColIndex(lo, "DataFim")).Value = dtFim
            .Cells(1, TableColIndex(lo, "Observacoes")).Value = obs
            .Cells(1, TableColIndex(lo, "DataRegistro")).Value = Now
            .Cells(1, TableColIndex(lo, "Usuario")).Value = Application.UserName
        End With
    Else
        Allocation_ValidateNoOverlapExcluding lo, funcionarioId, dtIni, dtFim, alocId
        Allocation_ValidateCapacityExcluding lo, regiaoCodigo, dtIni, dtFim, alocId

        With lo.DataBodyRange.Rows(rowIdx)
            .Cells(1, TableColIndex(lo, "FuncionarioID")).Value = funcionarioId
            .Cells(1, TableColIndex(lo, "RegiaoCodigo")).Value = regiaoCodigo
            .Cells(1, TableColIndex(lo, "DataInicio")).Value = dtIni
            .Cells(1, TableColIndex(lo, "DataFim")).Value = dtFim
            .Cells(1, TableColIndex(lo, "Observacoes")).Value = obs
            .Cells(1, TableColIndex(lo, "Usuario")).Value = Application.UserName
        End With
    End If

    wsDb.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True

    Dashboard_RefreshAll
    Dim msgOk As String
    ws.Unprotect Password:=pwd
    ws.Range("B2").Value = alocId
    ws.Protect Password:=pwd, UserInterfaceOnly:=True

    If rowIdx = 0 Then
        msgOk = "Alocacao salva para " & funcionarioId & " em " & regiaoCodigo
    Else
        msgOk = "Alocacao atualizada: " & alocId
    End If
    If autoClosed Then
        msgOk = msgOk & vbCrLf & "Alocacao anterior encerrada em " & Format$(autoNewFim, "dd/mm/yyyy")
    End If
    MsgBox msgOk, vbInformation, APP_TITLE
    Exit Sub
ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If autoClosed And Not wsDb Is Nothing Then
        wsDb.Unprotect Password:=pwd
        lo.DataBodyRange.Cells(autoRow, TableColIndex(lo, "DataFim")).Value = autoOldFim
    End If
    If Not wsDb Is Nothing Then
        wsDb.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
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

Private Sub Allocation_ValidateNoOverlapExcluding(ByVal lo As ListObject, ByVal funcionarioId As String, ByVal dtIni As Date, ByVal dtFim As Date, ByVal excludeAlocId As String)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim idxAid As Long
    Dim idxEmp As Long
    Dim idxIni As Long
    Dim idxFim As Long
    idxAid = TableColIndex(lo, "AlocacaoID")
    idxEmp = TableColIndex(lo, "FuncionarioID")
    idxIni = TableColIndex(lo, "DataInicio")
    idxFim = TableColIndex(lo, "DataFim")

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxAid).Value), excludeAlocId, vbTextCompare) <> 0 Then
            If StrComp(CStr(lo.DataBodyRange.Cells(r, idxEmp).Value), funcionarioId, vbTextCompare) = 0 Then
                If DateRangesOverlap(dtIni, dtFim, CDate(lo.DataBodyRange.Cells(r, idxIni).Value), CDate(lo.DataBodyRange.Cells(r, idxFim).Value)) Then
                    Err.Raise vbObjectError + 330, APP_TITLE, "Sobreposicao de periodos para o mesmo funcionario."
                End If
            End If
        End If
    Next r
End Sub

Private Sub Allocation_ValidateNoOverlapWithCandidate(ByVal lo As ListObject, ByVal funcionarioId As String, ByVal dtIni As Date, ByVal dtFim As Date, ByVal candidateRow As Long, ByVal hasCandidate As Boolean, ByVal candidateNewFim As Date)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim idxEmp As Long
    Dim idxIni As Long
    Dim idxFim As Long
    idxEmp = TableColIndex(lo, "FuncionarioID")
    idxIni = TableColIndex(lo, "DataInicio")
    idxFim = TableColIndex(lo, "DataFim")

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxEmp).Value), funcionarioId, vbTextCompare) = 0 Then
            Dim existingIni As Date
            Dim existingFim As Date
            existingIni = CDate(lo.DataBodyRange.Cells(r, idxIni).Value)
            existingFim = CDate(lo.DataBodyRange.Cells(r, idxFim).Value)

            If hasCandidate And r = candidateRow Then existingFim = candidateNewFim

            If existingFim >= existingIni Then
                If DateRangesOverlap(dtIni, dtFim, existingIni, existingFim) Then
                    Err.Raise vbObjectError + 330, APP_TITLE, "Sobreposicao de periodos para o mesmo funcionario."
                End If
            End If
        End If
    Next r
End Sub

Private Sub Allocation_ValidateCapacity(ByVal regiaoCodigo As String, ByVal dtIni As Date, ByVal dtFim As Date)
    Dim cap As Long
    cap = Region_GetCapacity(regiaoCodigo)
    If cap <= 0 Then Exit Sub

    If Allocation_MaxConcurrentForRegion(regiaoCodigo, dtIni, dtFim, vbNullString) > cap Then
        Err.Raise vbObjectError + 340, APP_TITLE, "Capacidade maxima excedida para o periodo informado."
    End If
End Sub

Private Sub Allocation_ValidateCapacityExcluding(ByVal lo As ListObject, ByVal regiaoCodigo As String, ByVal dtIni As Date, ByVal dtFim As Date, ByVal excludeAlocId As String)
    Dim cap As Long
    cap = Region_GetCapacity(regiaoCodigo)
    If cap <= 0 Then Exit Sub

    If Allocation_MaxConcurrentForRegion(regiaoCodigo, dtIni, dtFim, excludeAlocId) > cap Then
        Err.Raise vbObjectError + 340, APP_TITLE, "Capacidade maxima excedida para o periodo informado."
    End If
End Sub

Private Function Allocation_FindRowById(ByVal lo As ListObject, ByVal alocId As String) As Long
    Allocation_FindRowById = 0
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Len(Trim$(alocId)) = 0 Then Exit Function

    Dim idxAid As Long
    idxAid = TableColIndex(lo, "AlocacaoID")
    If idxAid = 0 Then Exit Function

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxAid).Value), alocId, vbTextCompare) = 0 Then
            Allocation_FindRowById = r
            Exit Function
        End If
    Next r
End Function

Public Sub Allocation_LoadToFormById(ByVal alocId As String)
    Dim wsDb As Worksheet
    Set wsDb = GetWs(SH_ALOC_DB)
    Dim lo As ListObject
    Set lo = wsDb.ListObjects(TB_ALOC)
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 360, APP_TITLE, "Nao ha alocacoes para carregar."

    Dim rowIdx As Long
    rowIdx = Allocation_FindRowById(lo, alocId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 361, APP_TITLE, "Alocacao nao encontrada: " & alocId

    Dim ws As Worksheet
    Set ws = GetWs(SH_ALOC_FORM)
    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Unprotect Password:=pwd

    ws.Range("B2").Value = alocId
    ws.Range("B3").Value = CStr(lo.DataBodyRange.Cells(rowIdx, TableColIndex(lo, "FuncionarioID")).Value)
    ws.Range("B4").Value = CStr(lo.DataBodyRange.Cells(rowIdx, TableColIndex(lo, "RegiaoCodigo")).Value)
    ws.Range("B5").Value = CDate(lo.DataBodyRange.Cells(rowIdx, TableColIndex(lo, "DataInicio")).Value)
    ws.Range("B6").Value = CDate(lo.DataBodyRange.Cells(rowIdx, TableColIndex(lo, "DataFim")).Value)
    ws.Range("B7").Value = CStr(lo.DataBodyRange.Cells(rowIdx, TableColIndex(lo, "Observacoes")).Value)
    ws.Range("B9").Value = "NAO"
    ws.Range("B10").ClearContents

    ws.Protect Password:=pwd, UserInterfaceOnly:=True
End Sub

Public Sub Allocation_DeleteById(ByVal alocId As String)
    Dim wsDb As Worksheet
    Set wsDb = GetWs(SH_ALOC_DB)
    Dim lo As ListObject
    Set lo = wsDb.ListObjects(TB_ALOC)
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 370, APP_TITLE, "Nao ha alocacoes para excluir."

    Dim rowIdx As Long
    rowIdx = Allocation_FindRowById(lo, alocId)
    If rowIdx = 0 Then Err.Raise vbObjectError + 371, APP_TITLE, "Alocacao nao encontrada: " & alocId

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsDb.Unprotect Password:=pwd
    lo.ListRows(rowIdx).Delete
    wsDb.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True

    Setup_RefreshAfterDataChange
End Sub

Private Function Allocation_FindAutoCloseCandidate(ByVal lo As ListObject, ByVal funcionarioId As String, ByVal dtIni As Date, ByVal dtFim As Date, ByRef rowIndex As Long, ByRef oldFim As Date, ByRef newFim As Date) As Boolean
    Allocation_FindAutoCloseCandidate = False
    rowIndex = 0

    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim idxEmp As Long
    Dim idxIni As Long
    Dim idxFim As Long
    idxEmp = TableColIndex(lo, "FuncionarioID")
    idxIni = TableColIndex(lo, "DataInicio")
    idxFim = TableColIndex(lo, "DataFim")
    If idxEmp = 0 Or idxIni = 0 Or idxFim = 0 Then Exit Function

    Dim r As Long
    Dim hits As Long
    hits = 0

    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxEmp).Value), funcionarioId, vbTextCompare) = 0 Then
            If IsDate(lo.DataBodyRange.Cells(r, idxIni).Value) And IsDate(lo.DataBodyRange.Cells(r, idxFim).Value) Then
                Dim oldIniLocal As Date
                Dim oldFimLocal As Date
                oldIniLocal = CDate(lo.DataBodyRange.Cells(r, idxIni).Value)
                oldFimLocal = CDate(lo.DataBodyRange.Cells(r, idxFim).Value)

                If DateRangesOverlap(dtIni, dtFim, oldIniLocal, oldFimLocal) Then
                    If dtIni <= oldIniLocal Then
                        Err.Raise vbObjectError + 331, APP_TITLE, "Existe uma alocacao que inicia em " & Format$(oldIniLocal, "dd/mm/yyyy") & ". Para realocar, use DataInicio maior que essa data ou ajuste a alocacao anterior."
                    End If
                    hits = hits + 1
                    rowIndex = r
                    oldFim = oldFimLocal
                End If
            End If
        End If
    Next r

    If hits = 0 Then Exit Function
    If hits > 1 Then Err.Raise vbObjectError + 332, APP_TITLE, "Foram encontradas multiplas alocacoes sobrepostas para este funcionario. Corrija a base antes de realocar."

    newFim = DateAdd("d", -1, dtIni)
    If newFim < CDate(lo.DataBodyRange.Cells(rowIndex, idxIni).Value) Then
        Err.Raise vbObjectError + 333, APP_TITLE, "Data de realocacao invalida para encerrar a alocacao anterior."
    End If

    Allocation_FindAutoCloseCandidate = True
End Function

Private Function Allocation_MaxConcurrentForRegion(ByVal regiaoCodigo As String, ByVal candidateIni As Date, ByVal candidateFim As Date, ByVal excludeAlocId As String) As Long
    Dim lo As ListObject
    Set lo = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)

    Dim idxAid As Long
    Dim idxReg As Long
    Dim idxIni As Long
    Dim idxFim As Long
    idxAid = TableColIndex(lo, "AlocacaoID")
    idxReg = TableColIndex(lo, "RegiaoCodigo")
    idxIni = TableColIndex(lo, "DataInicio")
    idxFim = TableColIndex(lo, "DataFim")

    Dim dates() As Date
    Dim countDates As Long
    countDates = 0
    Allocation_AddUniqueDate dates, countDates, candidateIni
    Allocation_AddUniqueDate dates, countDates, candidateFim

    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo.DataBodyRange.Rows.Count
            If Len(excludeAlocId) > 0 Then
                If StrComp(CStr(lo.DataBodyRange.Cells(r, idxAid).Value), excludeAlocId, vbTextCompare) = 0 Then GoTo NextRow
            End If

            If StrComp(CStr(lo.DataBodyRange.Cells(r, idxReg).Value), regiaoCodigo, vbTextCompare) = 0 Then
                Dim existingIni As Date
                Dim existingFim As Date
                existingIni = CDate(lo.DataBodyRange.Cells(r, idxIni).Value)
                existingFim = CDate(lo.DataBodyRange.Cells(r, idxFim).Value)

                If DateRangesOverlap(candidateIni, candidateFim, existingIni, existingFim) Then
                    If existingIni >= candidateIni And existingIni <= candidateFim Then Allocation_AddUniqueDate dates, countDates, existingIni
                    If existingFim >= candidateIni And existingFim <= candidateFim Then Allocation_AddUniqueDate dates, countDates, existingFim
                End If
            End If
NextRow:
        Next r
    End If

    Dim i As Long
    For i = 1 To countDates
        Dim currentCount As Long
        currentCount = 1

        If Not lo.DataBodyRange Is Nothing Then
            Dim rr As Long
            For rr = 1 To lo.DataBodyRange.Rows.Count
                If Len(excludeAlocId) > 0 Then
                    If StrComp(CStr(lo.DataBodyRange.Cells(rr, idxAid).Value), excludeAlocId, vbTextCompare) = 0 Then GoTo NextInner
                End If

                If StrComp(CStr(lo.DataBodyRange.Cells(rr, idxReg).Value), regiaoCodigo, vbTextCompare) = 0 Then
                    If dates(i) >= CDate(lo.DataBodyRange.Cells(rr, idxIni).Value) And dates(i) <= CDate(lo.DataBodyRange.Cells(rr, idxFim).Value) Then
                        currentCount = currentCount + 1
                    End If
                End If
NextInner:
            Next rr
        End If

        If currentCount > Allocation_MaxConcurrentForRegion Then
            Allocation_MaxConcurrentForRegion = currentCount
        End If
    Next i
End Function

Private Sub Allocation_AddUniqueDate(ByRef dates() As Date, ByRef countDates As Long, ByVal value As Date)
    Dim i As Long
    For i = 1 To countDates
        If dates(i) = value Then Exit Sub
    Next i

    countDates = countDates + 1
    If countDates = 1 Then
        ReDim dates(1 To 1)
    Else
        ReDim Preserve dates(1 To countDates)
    End If
    dates(countDates) = value
End Sub

