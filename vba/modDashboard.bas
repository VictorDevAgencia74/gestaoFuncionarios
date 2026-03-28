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

    Dim loD As ListObject
    Set loD = wsDash.ListObjects(TB_DASH)

    If Not loD.DataBodyRange Is Nothing Then loD.DataBodyRange.Delete

    If loR.DataBodyRange Is Nothing Then Exit Sub

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

    For r = 1 To loR.DataBodyRange.Rows.Count
        Dim lr As ListRow
        Set lr = loD.ListRows.Add
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

        lr.Range.Cells(1, 1).Value = regCode
        lr.Range.Cells(1, 2).Value = regName
        lr.Range.Cells(1, 3).Value = cap
        lr.Range.Cells(1, 4).Value = alocados
        If cap <= 0 Then
            lr.Range.Cells(1, 5).Value = 0
        Else
            lr.Range.Cells(1, 5).Value = alocados / cap
        End If
    Next r
    If Not loD.DataBodyRange Is Nothing Then
        loD.ListColumns("TaxaOcupacao").DataBodyRange.NumberFormat = "0.0%"
        loD.ListColumns("CapacidadeMaxima").DataBodyRange.NumberFormat = "0"
        loD.ListColumns("AlocadosHoje").DataBodyRange.NumberFormat = "0"
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
    If Not loA.DataBodyRange Is Nothing Then
        Dim idxFim As Long
        idxFim = TableColIndex(loA, "DataFim")
        Dim ra As Long
        For ra = 1 To loA.DataBodyRange.Rows.Count
            Dim df As Date
            df = CDate(loA.DataBodyRange.Cells(ra, idxFim).Value)
            If df >= Date And df <= Date + 7 Then vencendo = vencendo + 1
        Next ra
    End If
    wsDash.Range("B6").Value = vencendo
End Sub

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
    Set ptDest = wsDash.Range("H3")

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

    Set co = wsDash.ChartObjects.Add(Left:=wsDash.Range("H10").Left, Top:=wsDash.Range("H10").Top, Width:=420, Height:=240)
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

