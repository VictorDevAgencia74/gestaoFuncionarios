Attribute VB_Name = "modUtils"
Option Explicit

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
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
End Sub

Public Sub ApplySheetTheme(ByVal ws As Worksheet, ByVal titleText As String, ByVal titleRangeAddress As String)
    ws.Cells.WrapText = False
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.HorizontalAlignment = xlLeft
    ws.Rows.RowHeight = 18
    ws.Columns.ColumnWidth = 12

    ws.Range(titleRangeAddress).UnMerge
    ws.Range(titleRangeAddress).Merge
    ws.Range(titleRangeAddress).Value = titleText
    With ws.Range(titleRangeAddress)
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(33, 115, 70)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(ws.Range(titleRangeAddress).Row).RowHeight = 34

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
        lo.TableStyle = "TableStyleMedium2"
    Else
        For i = LBound(headers) To UBound(headers)
            lo.HeaderRowRange.Cells(1, 1 + (i - LBound(headers))).Value = headers(i)
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
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPt, topPt, widthPt, heightPt)
    btn.TextFrame2.TextRange.Text = caption
    btn.TextFrame2.TextRange.Font.Size = 11
    btn.TextFrame2.TextRange.Font.Bold = msoTrue
    btn.Fill.ForeColor.RGB = RGB(33, 115, 70)
    btn.Line.ForeColor.RGB = RGB(33, 115, 70)
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.OnAction = macroName
    btn.Placement = xlMoveAndSize
End Sub

Public Sub AddSheetButtonAtRange(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal area As Range)
    AddSheetButton ws, caption, macroName, area.Left, area.Top, area.Width, area.Height
End Sub

