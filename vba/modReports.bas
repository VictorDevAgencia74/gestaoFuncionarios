Attribute VB_Name = "modReports"
Option Explicit

Public Sub Reports_GenerateMonthlyPDFs(Optional ByVal yearV As Long = 0, Optional ByVal monthV As Long = 0)
    On Error GoTo ErrHandler

    Dim y As Long
    Dim m As Long
    y = yearV
    m = monthV
    If y = 0 Then y = CLng(InputBox("Ano (ex: 2026):", APP_TITLE, Year(Date)))
    If m = 0 Then m = CLng(InputBox("Mes (1-12):", APP_TITLE, Month(Date)))
    If y < 2000 Or y > 2100 Then Err.Raise vbObjectError + 500, APP_TITLE, "Ano invalido."
    If m < 1 Or m > 12 Then Err.Raise vbObjectError + 501, APP_TITLE, "Mes invalido."

    Dim dtIni As Date
    Dim dtFim As Date
    dtIni = DateSerial(y, m, 1)
    dtFim = DateSerial(y, m + 1, 0)

    Dim outFolder As String
    outFolder = EnsureFolder(EnsureFolder(WorkbookFolder() & "\\reports") & "\\" & Format$(y, "0000") & "-" & Format$(m, "00"))

    Dim ws As Worksheet
    Set ws = GetWs(SH_REL)
    ws.Visible = xlSheetVisible
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Cells.Clear

    Reports_Build ws, dtIni, dtFim

    Dim filePath As String
    filePath = outFolder & "\\Relatorio_" & Format$(y, "0000") & "-" & Format$(m, "00") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ws.Visible = xlSheetVeryHidden
    MsgBox "PDF gerado em: " & filePath, vbInformation, APP_TITLE
    Exit Sub

ErrHandler:
    On Error Resume Next
    GetWs(SH_REL).Visible = xlSheetVeryHidden
    MsgBox Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Reports_Build(ByVal ws As Worksheet, ByVal dtIni As Date, ByVal dtFim As Date)
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False

    ws.Range("A1").Value = "Relatorio Mensal de Alocacoes"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Value = "Periodo: " & Format$(dtIni, "dd/mm/yyyy") & " a " & Format$(dtFim, "dd/mm/yyyy")
    ws.Range("A3").Value = "Gerado em: " & Format$(Now, "dd/mm/yyyy hh:nn")

    Dim rowPtr As Long
    rowPtr = 5

    rowPtr = Reports_WriteAllocationsActiveAt(ws, rowPtr, dtFim)
    rowPtr = rowPtr + 2
    rowPtr = Reports_WriteMovements(ws, rowPtr, dtIni, dtFim)
    rowPtr = rowPtr + 2
    rowPtr = Reports_WriteOccupancy(ws, rowPtr, dtFim)

    ws.Columns.AutoFit
    ws.Range("A1:H1").EntireRow.RowHeight = 22
End Sub

Private Function Reports_WriteAllocationsActiveAt(ByVal ws As Worksheet, ByVal startRow As Long, ByVal refDate As Date) As Long
    ws.Range("A" & startRow).Value = "Alocacoes vigentes em " & Format$(refDate, "dd/mm/yyyy")
    ws.Range("A" & startRow).Font.Bold = True
    startRow = startRow + 1

    ws.Range("A" & startRow & ":H" & startRow).Value = Array("FuncionarioID", "Nome", "CPF", "Regiao", "DataInicio", "DataFim", "Supervisor", "Obs")
    ws.Range("A" & startRow & ":H" & startRow).Font.Bold = True
    startRow = startRow + 1

    Dim loA As ListObject
    Set loA = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)
    If loA.DataBodyRange Is Nothing Then
        Reports_WriteAllocationsActiveAt = startRow
        Exit Function
    End If

    Dim r As Long
    Dim idxEmp As Long, idxReg As Long, idxIni As Long, idxFim As Long, idxObs As Long
    idxEmp = TableColIndex(loA, "FuncionarioID")
    idxReg = TableColIndex(loA, "RegiaoCodigo")
    idxIni = TableColIndex(loA, "DataInicio")
    idxFim = TableColIndex(loA, "DataFim")
    idxObs = TableColIndex(loA, "Observacoes")

    For r = 1 To loA.DataBodyRange.Rows.Count
        Dim di As Date
        Dim df As Date
        di = CDate(loA.DataBodyRange.Cells(r, idxIni).Value)
        df = CDate(loA.DataBodyRange.Cells(r, idxFim).Value)
        If di <= refDate And df >= refDate Then
            Dim empId As String
            Dim regCode As String
            empId = CStr(loA.DataBodyRange.Cells(r, idxEmp).Value)
            regCode = CStr(loA.DataBodyRange.Cells(r, idxReg).Value)

            ws.Cells(startRow, 1).Value = empId
            ws.Cells(startRow, 2).Value = Employee_GetName(empId)
            ws.Cells(startRow, 3).Value = Employee_GetCPF(empId)
            ws.Cells(startRow, 4).Value = regCode & " - " & Region_GetName(regCode)
            ws.Cells(startRow, 5).Value = di
            ws.Cells(startRow, 6).Value = df
            ws.Cells(startRow, 7).Value = Region_GetSupervisor(regCode)
            ws.Cells(startRow, 8).Value = CStr(loA.DataBodyRange.Cells(r, idxObs).Value)

            startRow = startRow + 1
        End If
    Next r

    Reports_WriteAllocationsActiveAt = startRow
End Function

Private Function Reports_WriteMovements(ByVal ws As Worksheet, ByVal startRow As Long, ByVal dtIni As Date, ByVal dtFim As Date) As Long
    ws.Range("A" & startRow).Value = "Movimentacoes no periodo"
    ws.Range("A" & startRow).Font.Bold = True
    startRow = startRow + 1

    ws.Range("A" & startRow & ":G" & startRow).Value = Array("FuncionarioID", "Nome", "Regiao", "DataInicio", "DataFim", "Movimento", "Obs")
    ws.Range("A" & startRow & ":G" & startRow).Font.Bold = True
    startRow = startRow + 1

    Dim loA As ListObject
    Set loA = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)
    If loA.DataBodyRange Is Nothing Then
        Reports_WriteMovements = startRow
        Exit Function
    End If

    Dim r As Long
    Dim idxEmp As Long, idxReg As Long, idxIni As Long, idxFim As Long, idxObs As Long
    idxEmp = TableColIndex(loA, "FuncionarioID")
    idxReg = TableColIndex(loA, "RegiaoCodigo")
    idxIni = TableColIndex(loA, "DataInicio")
    idxFim = TableColIndex(loA, "DataFim")
    idxObs = TableColIndex(loA, "Observacoes")

    For r = 1 To loA.DataBodyRange.Rows.Count
        Dim di As Date
        Dim df As Date
        di = CDate(loA.DataBodyRange.Cells(r, idxIni).Value)
        df = CDate(loA.DataBodyRange.Cells(r, idxFim).Value)
        Dim moved As String
        moved = ""
        If di >= dtIni And di <= dtFim Then moved = "Inicio"
        If df >= dtIni And df <= dtFim Then
            If Len(moved) > 0 Then
                moved = moved & "+Fim"
            Else
                moved = "Fim"
            End If
        End If
        If Len(moved) = 0 Then GoTo NextR

        Dim empId As String
        Dim regCode As String
        empId = CStr(loA.DataBodyRange.Cells(r, idxEmp).Value)
        regCode = CStr(loA.DataBodyRange.Cells(r, idxReg).Value)

        ws.Cells(startRow, 1).Value = empId
        ws.Cells(startRow, 2).Value = Employee_GetName(empId)
        ws.Cells(startRow, 3).Value = regCode & " - " & Region_GetName(regCode)
        ws.Cells(startRow, 4).Value = di
        ws.Cells(startRow, 5).Value = df
        ws.Cells(startRow, 6).Value = moved
        ws.Cells(startRow, 7).Value = CStr(loA.DataBodyRange.Cells(r, idxObs).Value)
        startRow = startRow + 1

NextR:
    Next r

    Reports_WriteMovements = startRow
End Function

Private Function Reports_WriteOccupancy(ByVal ws As Worksheet, ByVal startRow As Long, ByVal refDate As Date) As Long
    ws.Range("A" & startRow).Value = "Ocupacao por regiao (vigente em " & Format$(refDate, "dd/mm/yyyy") & ")"
    ws.Range("A" & startRow).Font.Bold = True
    startRow = startRow + 1

    ws.Range("A" & startRow & ":E" & startRow).Value = Array("RegiaoCodigo", "RegiaoNome", "Capacidade", "Alocados", "Taxa")
    ws.Range("A" & startRow & ":E" & startRow).Font.Bold = True
    startRow = startRow + 1

    Dim loR As ListObject
    Set loR = GetWs(SH_REGIOES).ListObjects(TB_REG)
    If loR.DataBodyRange Is Nothing Then
        Reports_WriteOccupancy = startRow
        Exit Function
    End If

    Dim loA As ListObject
    Set loA = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)

    Dim r As Long
    Dim idxCode As Long, idxName As Long, idxCap As Long
    idxCode = TableColIndex(loR, "RegiaoCodigo")
    idxName = TableColIndex(loR, "RegiaoNome")
    idxCap = TableColIndex(loR, "CapacidadeMaxima")
    For r = 1 To loR.DataBodyRange.Rows.Count
        Dim regCode As String
        Dim cap As Long
        regCode = CStr(loR.DataBodyRange.Cells(r, idxCode).Value)
        cap = CLng(loR.DataBodyRange.Cells(r, idxCap).Value)
        Dim alocados As Long
        alocados = Reports_CountActiveInRegionAt(loA, regCode, refDate)

        ws.Cells(startRow, 1).Value = regCode
        ws.Cells(startRow, 2).Value = CStr(loR.DataBodyRange.Cells(r, idxName).Value)
        ws.Cells(startRow, 3).Value = cap
        ws.Cells(startRow, 4).Value = alocados
        If cap = 0 Then
            ws.Cells(startRow, 5).Value = 0
        Else
            ws.Cells(startRow, 5).Value = alocados / cap
        End If
        ws.Cells(startRow, 5).NumberFormat = "0.0%"
        startRow = startRow + 1
    Next r

    Reports_WriteOccupancy = startRow
End Function

Private Function Reports_CountActiveInRegionAt(ByVal loA As ListObject, ByVal regCode As String, ByVal refDate As Date) As Long
    Reports_CountActiveInRegionAt = 0
    If loA.DataBodyRange Is Nothing Then Exit Function
    Dim idxReg As Long, idxIni As Long, idxFim As Long
    idxReg = TableColIndex(loA, "RegiaoCodigo")
    idxIni = TableColIndex(loA, "DataInicio")
    idxFim = TableColIndex(loA, "DataFim")

    Dim r As Long
    For r = 1 To loA.DataBodyRange.Rows.Count
        If StrComp(CStr(loA.DataBodyRange.Cells(r, idxReg).Value), regCode, vbTextCompare) = 0 Then
            Dim di As Date
            Dim df As Date
            di = CDate(loA.DataBodyRange.Cells(r, idxIni).Value)
            df = CDate(loA.DataBodyRange.Cells(r, idxFim).Value)
            If di <= refDate And df >= refDate Then Reports_CountActiveInRegionAt = Reports_CountActiveInRegionAt + 1
        End If
    Next r
End Function

