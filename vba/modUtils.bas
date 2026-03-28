Attribute VB_Name = "modUtils"
Option Explicit

Public Function UI_FontBase() As String
    UI_FontBase = "Segoe UI"
End Function

Public Function UI_ColorPrimary() As Long
    UI_ColorPrimary = RGB(37, 99, 235)
End Function

Public Function UI_ColorSurface() As Long
    UI_ColorSurface = RGB(255, 255, 255)
End Function

Public Function UI_ColorSurfaceAlt() As Long
    UI_ColorSurfaceAlt = RGB(248, 250, 252)
End Function

Public Function UI_ColorBorder() As Long
    UI_ColorBorder = RGB(226, 232, 240)
End Function

Public Function UI_ColorText() As Long
    UI_ColorText = RGB(15, 23, 42)
End Function

Public Function UI_ColorTextMuted() As Long
    UI_ColorTextMuted = RGB(71, 85, 105)
End Function

Public Function UI_ColorTextOnPrimary() As Long
    UI_ColorTextOnPrimary = RGB(255, 255, 255)
End Function

Public Sub UI_StyleInputs(ByVal area As Range)
    area.Font.Name = UI_FontBase()
    area.Font.Color = UI_ColorText()
    area.Interior.Color = UI_ColorSurfaceAlt()
    area.Borders.LineStyle = xlContinuous
    area.Borders.Color = UI_ColorBorder()
End Sub

Public Sub UI_StyleLabels(ByVal area As Range)
    area.Font.Name = UI_FontBase()
    area.Font.Color = UI_ColorText()
    area.Font.Bold = True
End Sub

Public Sub UI_StyleKpi(ByVal area As Range)
    area.Font.Name = UI_FontBase()
    area.Font.Color = UI_ColorText()
    area.Font.Bold = True
    area.Interior.Color = UI_ColorSurfaceAlt()
    area.Borders.LineStyle = xlContinuous
    area.Borders.Color = UI_ColorBorder()
    area.HorizontalAlignment = xlCenter
    area.VerticalAlignment = xlCenter
End Sub

Public Function GetWs(ByVal sheetName As String) As Worksheet
    Set GetWs = ThisWorkbook.Worksheets(sheetName)
End Function

Public Function EnsureWorksheet(ByVal sheetName As String, Optional ByVal afterSheet As Variant) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        If IsMissing(afterSheet) Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
        End If
        ws.Name = sheetName
    End If
    Set EnsureWorksheet = ws
End Function

Public Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not (ws Is Nothing)
End Function

Public Function TableExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    TableExists = Not (lo Is Nothing)
End Function

Public Sub ClearSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells.NumberFormat = "General"
    ws.Cells.Font.Name = UI_FontBase()
    ws.Cells.Font.Size = 11
    ws.Cells.Font.Color = UI_ColorText()
    ws.Cells.Interior.Color = UI_ColorSurface()
End Sub

Public Sub ApplySheetTheme(ByVal ws As Worksheet, ByVal titleText As String, ByVal titleRangeAddress As String)
    ws.Cells.WrapText = False
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.HorizontalAlignment = xlLeft
    ws.Rows.RowHeight = 18
    ws.Columns.ColumnWidth = 12
    ws.Cells.Font.Name = UI_FontBase()
    ws.Cells.Font.Size = 11
    ws.Cells.Font.Color = UI_ColorText()
    ws.Cells.Interior.Color = UI_ColorSurface()

    ws.Range(titleRangeAddress).UnMerge
    ws.Range(titleRangeAddress).Merge
    ws.Range(titleRangeAddress).Value = titleText
    With ws.Range(titleRangeAddress)
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = UI_ColorTextOnPrimary()
        .Interior.Color = UI_ColorPrimary()
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(ws.Range(titleRangeAddress).Row).RowHeight = 34
    ws.Range(titleRangeAddress).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(titleRangeAddress).Borders(xlEdgeBottom).Color = UI_ColorBorder()

    Dim prev As Worksheet
    On Error Resume Next
    Set prev = ActiveSheet
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 110
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    If Not prev Is Nothing Then prev.Activate
    On Error GoTo 0
End Sub

Public Sub RemoveShapesByOnAction(ByVal ws As Worksheet, ParamArray macroNames() As Variant)
    Dim shp As Shape
    Dim i As Long
    For Each shp In ws.Shapes
        If Len(shp.OnAction) > 0 Then
            For i = LBound(macroNames) To UBound(macroNames)
                If StrComp(shp.OnAction, CStr(macroNames(i)), vbTextCompare) = 0 Then
                    shp.Delete
                    Exit For
                End If
            Next i
        End If
    Next shp

    Dim btn As Object
    On Error Resume Next
    For Each btn In ws.Buttons
        For i = LBound(macroNames) To UBound(macroNames)
            If StrComp(CStr(btn.OnAction), CStr(macroNames(i)), vbTextCompare) = 0 Then
                btn.Delete
                Exit For
            End If
            Err.Clear
            Dim n As String
            n = CStr(btn.Name)
            If Err.Number = 0 Then
                If Left$(n, 4) = "btn_" Then
                    If InStr(1, n, "btn_" & CStr(macroNames(i)), vbTextCompare) = 1 Then
                        btn.Delete
                        Exit For
                    End If
                End If
            End If
        Next i
    Next btn
    On Error GoTo 0
End Sub

Public Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headerRow As Long, ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    Dim lastCol As Long
    Dim i As Long

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        lastCol = UBound(headers) - LBound(headers) + 1
        For i = LBound(headers) To UBound(headers)
            ws.Cells(headerRow, 1 + (i - LBound(headers))).Value = headers(i)
            ws.Cells(headerRow, 1 + (i - LBound(headers))).Font.Bold = True
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow + 1, lastCol)), , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleMedium9"
    Else
        lastCol = UBound(headers) - LBound(headers) + 1

        On Error Resume Next
        lo.TableStyle = "TableStyleMedium9"
        On Error GoTo 0

        If lo.ListColumns.Count < lastCol Then
            For i = lo.ListColumns.Count + 1 To lastCol
                lo.ListColumns.Add
            Next i
        End If

        For i = 1 To lastCol
            On Error Resume Next
            lo.ListColumns(i).Name = CStr(headers(LBound(headers) + (i - 1)))
            On Error GoTo 0
        Next i
    End If

    Set EnsureTable = lo
End Function

Public Function TableColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(CStr(lc.Name), colName, vbTextCompare) = 0 Then
            TableColIndex = lc.Index
            Exit Function
        End If
    Next lc
    TableColIndex = 0
End Function

Public Function LastDataRow(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then
        LastDataRow = lo.HeaderRowRange.Row
    Else
        LastDataRow = lo.DataBodyRange.Rows(lo.DataBodyRange.Rows.Count).Row
    End If
End Function

Public Function NormalizeDigits(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    NormalizeDigits = out
End Function

Public Function IsValidCPF(ByVal cpf As String) As Boolean
    Dim d As String
    d = NormalizeDigits(cpf)
    If Len(d) <> 11 Then Exit Function
    If d = String$(11, Left$(d, 1)) Then Exit Function

    Dim i As Long
    Dim sum As Long
    Dim modv As Long
    Dim dig1 As Long
    Dim dig2 As Long

    sum = 0
    For i = 1 To 9
        sum = sum + CLng(Mid$(d, i, 1)) * (11 - i)
    Next i
    modv = (sum * 10) Mod 11
    If modv = 10 Then modv = 0
    dig1 = modv

    sum = 0
    For i = 1 To 10
        sum = sum + CLng(Mid$(d, i, 1)) * (12 - i)
    Next i
    modv = (sum * 10) Mod 11
    If modv = 10 Then modv = 0
    dig2 = modv

    IsValidCPF = (dig1 = CLng(Mid$(d, 10, 1)) And dig2 = CLng(Mid$(d, 11, 1)))
End Function

Public Function DateRangesOverlap(ByVal aStart As Date, ByVal aEnd As Date, ByVal bStart As Date, ByVal bEnd As Date) As Boolean
    DateRangesOverlap = (aStart <= bEnd) And (bStart <= aEnd)
End Function

Public Function GetConfigValue(ByVal address As String) As Variant
    GetConfigValue = GetWs(SH_CONFIG).Range(address).Value
End Function

Public Sub SetConfigValue(ByVal address As String, ByVal value As Variant)
    GetWs(SH_CONFIG).Range(address).Value = value
End Sub

Public Function EnsureFolder(ByVal folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    EnsureFolder = folderPath
End Function

Public Function WorkbookFolder() As String
    If Len(ThisWorkbook.Path) = 0 Then
        WorkbookFolder = CurDir$
    Else
        WorkbookFolder = ThisWorkbook.Path
    End If
End Function

Public Function NewGuidId() As String
    NewGuidId = Mid$(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
End Function

Public Sub AddSheetButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal leftPt As Double, ByVal topPt As Double, ByVal widthPt As Double, ByVal heightPt As Double)
    Dim shp As Object

    On Error Resume Next
    Dim b As Object
    For Each b In ws.Buttons
        If StrComp(CStr(b.OnAction), macroName, vbTextCompare) = 0 Then b.Delete
    Next b
    On Error GoTo 0

    Dim padX As Double
    Dim padY As Double
    padX = 6
    padY = 4

    If widthPt < (padX * 2 + 10) Then padX = 2
    If heightPt < (padY * 2 + 10) Then padY = 2

    Set shp = ws.Shapes.AddShape(5, leftPt + padX, topPt + padY, widthPt - (padX * 2), heightPt - (padY * 2))

    On Error Resume Next
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Name = UI_FontBase()
    shp.TextFrame.Characters.Font.Size = 11
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Color = UI_ColorTextOnPrimary()
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.Fill.ForeColor.RGB = UI_ColorPrimary()
    shp.Line.ForeColor.RGB = UI_ColorPrimary()
    shp.Shadow.Visible = False
    shp.Adjustments.Item(1) = 0.18
    shp.OnAction = macroName
    shp.Placement = xlMoveAndSize
    On Error GoTo 0
End Sub

Public Sub AddSheetButtonAtRange(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal area As Range)
    AddSheetButton ws, caption, macroName, area.Left, area.Top, area.Width, area.Height
End Sub

