Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub Dashboard_RefreshAll()
    On Error GoTo ErrHandler

    If Not WorksheetExists(SH_DASH) Then Exit Sub
    If Not WorksheetExists(SH_REGIOES) Then Exit Sub
    If Not WorksheetExists(SH_FUNC_DB) Then Exit Sub
    If Not WorksheetExists(SH_ALOC_DB) Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetWs(SH_DASH)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))

    Dashboard_RebuildActiveAllocationsTable ws, Date
    Dashboard_RebuildRegionSummary ws
    Dashboard_UpdateIndicators ws
    Dashboard_RebuildExpiryList ws
    Dashboard_RebuildPivot ws

    ws.Columns.AutoFit
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    If errNum <> 0 Or Len(errDesc) > 0 Then
        If Len(errDesc) = 0 Then errDesc = "Erro " & CStr(errNum)
        MsgBox "Dashboard: " & errDesc, vbExclamation, APP_TITLE
    End If
End Sub

Private Sub Dashboard_RebuildActiveAllocationsTable(ByVal wsDash As Worksheet, ByVal refDate As Date)
    Dim startCell As Range
    Set startCell = wsDash.Range("H20")
    startCell.Resize(1, 4).Value = Array("FuncionarioID", "RegiaoCodigo", "DataInicio", "DataFim")
    startCell.Resize(1, 4).Font.Bold = True

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsDash.ListObjects(TB_ALOC_HOJE)
    On Error GoTo 0
    If Not lo Is Nothing Then
        lo.Unlist
    End If

    Dim wsA As Worksheet
    Set wsA = GetWs(SH_ALOC_DB)
    Dim loA As ListObject
    Set loA = wsA.ListObjects(TB_ALOC)

    Dim outRow As Long
    outRow = startCell.Row + 1
    wsDash.Range(startCell.Offset(1, 0), startCell.Offset(2000, 3)).ClearContents

    If Not loA.DataBodyRange Is Nothing Then
        Dim r As Long
        Dim idxEmp As Long, idxReg As Long, idxIni As Long, idxFim As Long
        idxEmp = TableColIndex(loA, "FuncionarioID")
        idxReg = TableColIndex(loA, "RegiaoCodigo")
        idxIni = TableColIndex(loA, "DataInicio")
        idxFim = TableColIndex(loA, "DataFim")

        For r = 1 To loA.DataBodyRange.Rows.Count
            Dim di As Date
            Dim df As Date
            di = CDate(loA.DataBodyRange.Cells(r, idxIni).Value)
            df = CDate(loA.DataBodyRange.Cells(r, idxFim).Value)
            If (di <= refDate) And (df >= refDate) Then
                wsDash.Cells(outRow, startCell.Column).Value = CStr(loA.DataBodyRange.Cells(r, idxEmp).Value)
                wsDash.Cells(outRow, startCell.Column + 1).Value = CStr(loA.DataBodyRange.Cells(r, idxReg).Value)
                wsDash.Cells(outRow, startCell.Column + 2).Value = di
                wsDash.Cells(outRow, startCell.Column + 3).Value = df
                outRow = outRow + 1
            End If
        Next r
    End If

    Dim lastRow As Long
    Dim src As Range
    If outRow = startCell.Row + 1 Then
        Set src = startCell.Resize(1, 4)
    Else
        lastRow = outRow - 1
        Set src = wsDash.Range(startCell, wsDash.Cells(lastRow, startCell.Column + 3))
    End If
    Dim loNew As ListObject
    Set loNew = wsDash.ListObjects.Add(xlSrcRange, src, , xlYes)
    loNew.Name = TB_ALOC_HOJE
    loNew.TableStyle = "TableStyleMedium9"
    If Not loNew.DataBodyRange Is Nothing Then
        loNew.ListColumns("DataInicio").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        loNew.ListColumns("DataFim").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    End If
End Sub

Private Sub Dashboard_RebuildRegionSummary(ByVal wsDash As Worksheet)
    Dim wsR As Worksheet
    Set wsR = GetWs(SH_REGIOES)
    Dim loR As ListObject
    Set loR = wsR.ListObjects(TB_REG)

    Dim startCell As Range
    Set startCell = wsDash.Range("A9")
    startCell.Resize(1, 5).Value = Array("RegiaoCodigo", "RegiaoNome", "CapacidadeMaxima", "AlocadosHoje", "TaxaOcupacao")
    startCell.Resize(1, 5).Font.Bold = True

    Dim loD As ListObject
    On Error Resume Next
    Set loD = wsDash.ListObjects(TB_DASH)
    On Error GoTo 0

    Dim clearRows As Long
    clearRows = 1
    If Not loR.DataBodyRange Is Nothing Then clearRows = 1 + loR.DataBodyRange.Rows.Count
    If Not loD Is Nothing Then
        If loD.Range.Rows.Count > clearRows Then clearRows = loD.Range.Rows.Count
        loD.Unlist
    End If

    wsDash.Range(startCell.Offset(1, 0), startCell.Offset(clearRows, 4)).ClearContents
    wsDash.Range(startCell.Offset(1, 0), startCell.Offset(clearRows, 4)).Interior.Pattern = xlNone

    If loR.DataBodyRange Is Nothing Then
        Dim loEmpty As ListObject
        Set loEmpty = wsDash.ListObjects.Add(xlSrcRange, startCell.Resize(1, 5), , xlYes)
        loEmpty.Name = TB_DASH
        loEmpty.TableStyle = "TableStyleMedium2"
        Exit Sub
    End If

    Dim dictCount As Object
    Set dictCount = CreateObject("Scripting.Dictionary")

    Dim loHoje As ListObject
    On Error Resume Next
    Set loHoje = wsDash.ListObjects(TB_ALOC_HOJE)
    On Error GoTo 0
    If Not loHoje Is Nothing Then
        If Not loHoje.DataBodyRange Is Nothing Then
            Dim rr As Long
            For rr = 1 To loHoje.DataBodyRange.Rows.Count
                Dim regCodeHoje As String
                regCodeHoje = CStr(loHoje.DataBodyRange.Cells(rr, 2).Value)
                If Len(regCodeHoje) > 0 Then
                    If dictCount.Exists(regCodeHoje) Then
                        dictCount(regCodeHoje) = CLng(dictCount(regCodeHoje)) + 1
                    Else
                        dictCount.Add regCodeHoje, 1
                    End If
                End If
            Next rr
        End If
    End If

    Dim r As Long
    Dim idxCode As Long
    Dim idxName As Long
    Dim idxCap As Long
    idxCode = TableColIndex(loR, "RegiaoCodigo")
    idxName = TableColIndex(loR, "RegiaoNome")
    idxCap = TableColIndex(loR, "CapacidadeMaxima")

    Dim outRow As Long
    outRow = startCell.Row + 1

    For r = 1 To loR.DataBodyRange.Rows.Count
        Dim regCode As String
        Dim regName As String
        Dim cap As Long
        Dim alocados As Long
        regCode = CStr(loR.DataBodyRange.Cells(r, idxCode).Value)
        regName = CStr(loR.DataBodyRange.Cells(r, idxName).Value)
        If IsNumeric(loR.DataBodyRange.Cells(r, idxCap).Value) Then
            cap = CLng(loR.DataBodyRange.Cells(r, idxCap).Value)
        Else
            cap = 0
        End If

        If dictCount.Exists(regCode) Then
            alocados = CLng(dictCount(regCode))
        Else
            alocados = 0
        End If

        wsDash.Cells(outRow, startCell.Column).Value = regCode
        wsDash.Cells(outRow, startCell.Column + 1).Value = regName
        wsDash.Cells(outRow, startCell.Column + 2).Value = cap
        wsDash.Cells(outRow, startCell.Column + 3).Value = alocados
        If cap <= 0 Then
            wsDash.Cells(outRow, startCell.Column + 4).Value = 0
        Else
            wsDash.Cells(outRow, startCell.Column + 4).Value = alocados / cap
        End If
        outRow = outRow + 1
    Next r

    Dim lastRow As Long
    Dim src As Range
    lastRow = outRow - 1
    Set src = wsDash.Range(startCell, wsDash.Cells(lastRow, startCell.Column + 4))

    Dim loNew As ListObject
    Set loNew = wsDash.ListObjects.Add(xlSrcRange, src, , xlYes)
    loNew.Name = TB_DASH
    loNew.TableStyle = "TableStyleMedium2"
    If Not loNew.DataBodyRange Is Nothing Then
        loNew.ListColumns("TaxaOcupacao").DataBodyRange.NumberFormat = "0.0%"
        loNew.ListColumns("CapacidadeMaxima").DataBodyRange.NumberFormat = "0"
        loNew.ListColumns("AlocadosHoje").DataBodyRange.NumberFormat = "0"
    End If
End Sub

Private Sub Dashboard_UpdateIndicators(ByVal wsDash As Worksheet)
    Dim wsF As Worksheet
    Set wsF = GetWs(SH_FUNC_DB)
    Dim loF As ListObject
    Set loF = wsF.ListObjects(TB_FUNC)

    Dim loHoje As ListObject
    Set loHoje = wsDash.ListObjects(TB_ALOC_HOJE)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Not loHoje.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To loHoje.DataBodyRange.Rows.Count
            dict(CStr(loHoje.DataBodyRange.Cells(r, 1).Value)) = True
        Next r
    End If

    Dim semAloc As Long
    semAloc = 0

    If Not loF.DataBodyRange Is Nothing Then
        Dim idxId As Long
        Dim idxStatus As Long
        idxId = TableColIndex(loF, "FuncionarioID")
        idxStatus = TableColIndex(loF, "Status")

        Dim rf As Long
        For rf = 1 To loF.DataBodyRange.Rows.Count
            If StrComp(CStr(loF.DataBodyRange.Cells(rf, idxStatus).Value), "Ativo", vbTextCompare) = 0 Then
                If Not dict.Exists(CStr(loF.DataBodyRange.Cells(rf, idxId).Value)) Then semAloc = semAloc + 1
            End If
        Next rf
    End If

    wsDash.Range("B5").Value = semAloc

    Dim vencendo As Long
    vencendo = 0
    Dim loA As ListObject
    Set loA = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)
    Dim warnDays As Long
    warnDays = Dashboard_GetWarnDays()
    If Not loA.DataBodyRange Is Nothing Then
        Dim idxIni As Long
        Dim idxFim As Long
        idxIni = TableColIndex(loA, "DataInicio")
        idxFim = TableColIndex(loA, "DataFim")
        Dim ra As Long
        For ra = 1 To loA.DataBodyRange.Rows.Count
            Dim di As Date
            Dim df As Date
            di = CDate(loA.DataBodyRange.Cells(ra, idxIni).Value)
            df = CDate(loA.DataBodyRange.Cells(ra, idxFim).Value)
            If di <= Date And df >= Date And df <= Date + warnDays Then vencendo = vencendo + 1
        Next ra
    End If
    wsDash.Range("B6").Value = vencendo
    wsDash.Range("A6").Value = "Alocacoes vencendo (" & CStr(warnDays) & " dias)"
End Sub

Private Sub Dashboard_RebuildExpiryList(ByVal wsDash As Worksheet)
    Dim warnDays As Long
    warnDays = Dashboard_GetWarnDays()

    Dim loRegions As ListObject
    Set loRegions = wsDash.ListObjects(TB_DASH)

    Dim startRow As Long
    startRow = 20
    On Error Resume Next
    startRow = LastDataRow(loRegions) + 3
    On Error GoTo 0
    If startRow < 20 Then startRow = 20

    Dim startCell As Range
    Set startCell = wsDash.Cells(startRow, 1)
    startCell.Resize(1, 6).Value = Array("FuncionarioID", "NomeCompleto", "RegiaoCodigo", "DataFim", "Situacao", "Dias")
    startCell.Resize(1, 6).Font.Bold = True

    Dim loOld As ListObject
    On Error Resume Next
    Set loOld = wsDash.ListObjects(TB_VENC)
    On Error GoTo 0
    If Not loOld Is Nothing Then
        loOld.Unlist
    End If

    wsDash.Range(startCell.Offset(1, 0), startCell.Offset(2000, 5)).ClearContents
    wsDash.Range(startCell.Offset(0, 0), startCell.Offset(2000, 5)).Interior.Pattern = xlNone

    Dim loA As ListObject
    Set loA = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)
    If loA.DataBodyRange Is Nothing Then Exit Sub

    Dim idxEmp As Long, idxReg As Long, idxIni As Long, idxFim As Long
    idxEmp = TableColIndex(loA, "FuncionarioID")
    idxReg = TableColIndex(loA, "RegiaoCodigo")
    idxIni = TableColIndex(loA, "DataInicio")
    idxFim = TableColIndex(loA, "DataFim")
    If idxEmp = 0 Or idxReg = 0 Or idxIni = 0 Or idxFim = 0 Then Exit Sub

    Dim dictCurrent As Object
    Dim dictLast As Object
    Set dictCurrent = CreateObject("Scripting.Dictionary")
    Set dictLast = CreateObject("Scripting.Dictionary")

    Dim today As Date
    today = Date

    Dim r As Long
    For r = 1 To loA.DataBodyRange.Rows.Count
        Dim empId As String
        Dim regCode As String
        Dim di As Date
        Dim df As Date

        empId = CStr(loA.DataBodyRange.Cells(r, idxEmp).Value)
        If Len(empId) = 0 Then GoTo NextRow
        regCode = CStr(loA.DataBodyRange.Cells(r, idxReg).Value)
        di = CDate(loA.DataBodyRange.Cells(r, idxIni).Value)
        df = CDate(loA.DataBodyRange.Cells(r, idxFim).Value)

        If Not dictLast.Exists(empId) Then
            dictLast.Add empId, Array(regCode, df)
        Else
            If df > CDate(dictLast(empId)(1)) Then dictLast(empId) = Array(regCode, df)
        End If

        If di <= today And df >= today Then
            If Not dictCurrent.Exists(empId) Then
                dictCurrent.Add empId, Array(regCode, df)
            Else
                If df > CDate(dictCurrent(empId)(1)) Then dictCurrent(empId) = Array(regCode, df)
            End If
        End If

NextRow:
    Next r

    Dim outRow As Long
    outRow = startCell.Row + 1

    Dim k As Variant
    For Each k In dictCurrent.Keys
        Dim endDt As Date
        endDt = CDate(dictCurrent(k)(1))
        If endDt <= today + warnDays Then
            wsDash.Cells(outRow, startCell.Column).Value = CStr(k)
            wsDash.Cells(outRow, startCell.Column + 1).Value = Employee_GetName(CStr(k))
            wsDash.Cells(outRow, startCell.Column + 2).Value = CStr(dictCurrent(k)(0))
            wsDash.Cells(outRow, startCell.Column + 3).Value = endDt
            wsDash.Cells(outRow, startCell.Column + 4).Value = "VENCENDO"
            wsDash.Cells(outRow, startCell.Column + 5).Value = CLng(endDt - today)
            outRow = outRow + 1
        End If
    Next k

    For Each k In dictLast.Keys
        If Not dictCurrent.Exists(k) Then
            Dim lastEnd As Date
            lastEnd = CDate(dictLast(k)(1))
            If lastEnd < today Then
                wsDash.Cells(outRow, startCell.Column).Value = CStr(k)
                wsDash.Cells(outRow, startCell.Column + 1).Value = Employee_GetName(CStr(k))
                wsDash.Cells(outRow, startCell.Column + 2).Value = CStr(dictLast(k)(0))
                wsDash.Cells(outRow, startCell.Column + 3).Value = lastEnd
                wsDash.Cells(outRow, startCell.Column + 4).Value = "VENCIDO"
                wsDash.Cells(outRow, startCell.Column + 5).Value = CLng(lastEnd - today)
                outRow = outRow + 1
            End If
        End If
    Next k

    Dim lastRow As Long
    Dim src As Range
    If outRow = startCell.Row + 1 Then
        Set src = startCell.Resize(1, 6)
    Else
        lastRow = outRow - 1
        Set src = wsDash.Range(startCell, wsDash.Cells(lastRow, startCell.Column + 5))
    End If

    Dim loNew As ListObject
    Set loNew = wsDash.ListObjects.Add(xlSrcRange, src, , xlYes)
    loNew.Name = TB_VENC
    loNew.TableStyle = "TableStyleMedium2"
    If Not loNew.DataBodyRange Is Nothing Then
        loNew.ListColumns("DataFim").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        loNew.ListColumns("Dias").DataBodyRange.NumberFormat = "0"
        Dashboard_SortExpiryTable loNew
        Dashboard_ApplyExpiryColors loNew
    End If
End Sub

Private Sub Dashboard_SortExpiryTable(ByVal lo As ListObject)
    On Error Resume Next
    lo.Sort.SortFields.Clear
    lo.Sort.SortFields.Add Key:=lo.ListColumns("DataFim").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    lo.Sort.Header = xlYes
    lo.Sort.Apply
    On Error GoTo 0
End Sub

Private Sub Dashboard_ApplyExpiryColors(ByVal lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim idxSit As Long
    idxSit = TableColIndex(lo, "Situacao")
    If idxSit = 0 Then Exit Sub

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim sit As String
        sit = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, idxSit).Value)))
        If sit = "VENCIDO" Then
            lo.DataBodyRange.Rows(r).Interior.Color = RGB(255, 199, 206)
        ElseIf sit = "VENCENDO" Then
            lo.DataBodyRange.Rows(r).Interior.Color = RGB(255, 235, 156)
        Else
            lo.DataBodyRange.Rows(r).Interior.Pattern = xlNone
        End If
    Next r
End Sub

Private Function Dashboard_GetWarnDays() As Long
    Dim v As Variant
    On Error Resume Next
    v = GetConfigValue(CFG_EXPIRY_WARN_DAYS_CELL)
    On Error GoTo 0
    If IsNumeric(v) Then
        Dashboard_GetWarnDays = CLng(v)
    Else
        Dashboard_GetWarnDays = 7
    End If
    If Dashboard_GetWarnDays < 0 Then Dashboard_GetWarnDays = 0
End Function

Private Sub Dashboard_RebuildPivot(ByVal wsDash As Worksheet)
    Dim loHoje As ListObject
    Set loHoje = wsDash.ListObjects(TB_ALOC_HOJE)

    Dim pt As PivotTable
    On Error Resume Next
    Set pt = wsDash.PivotTables("pvtAlocadosPorRegiao")
    If Not pt Is Nothing Then
        pt.TableRange2.Clear
    End If
    On Error GoTo 0

    If loHoje.DataBodyRange Is Nothing Then Exit Sub
    If loHoje.Range.Rows.Count < 2 Then Exit Sub

    Dim pc As PivotCache
    On Error Resume Next
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, loHoje.Range)
    On Error GoTo 0
    If pc Is Nothing Then Exit Sub

    Dim ptDest As Range
    Set ptDest = wsDash.Range("M3")

    On Error Resume Next
    Set pt = pc.CreatePivotTable(TableDestination:=ptDest, TableName:="pvtAlocadosPorRegiao")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    On Error Resume Next
    With pt
        .PivotFields("RegiaoCodigo").Orientation = xlRowField
        .AddDataField .PivotFields("FuncionarioID"), "Alocados", xlCount
        .RowAxisLayout xlTabularRow
    End With
    On Error GoTo 0

    Dim co As ChartObject
    On Error Resume Next
    Set co = wsDash.ChartObjects("chtAlocadosPorRegiao")
    On Error GoTo 0
    If Not co Is Nothing Then co.Delete

    Set co = wsDash.ChartObjects.Add(Left:=wsDash.Range("M10").Left, Top:=wsDash.Range("M10").Top, Width:=420, Height:=240)
    co.Name = "chtAlocadosPorRegiao"
    co.Chart.ChartType = xlColumnClustered
    If Not pt Is Nothing Then
        On Error Resume Next
        co.Chart.SetSourceData Source:=pt.TableRange1
        On Error GoTo 0
    End If
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Alocados por regiao (hoje)"
End Sub

