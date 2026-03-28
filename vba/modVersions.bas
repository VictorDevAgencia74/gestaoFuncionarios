Attribute VB_Name = "modVersions"
Option Explicit

Public Sub Version_AddEntry()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetWs(SH_VERSOES)

    Dim lo As ListObject
    Set lo = ws.ListObjects(TB_VERSOES)

    Dim curVersion As String
    curVersion = Version_GetCurrent(lo)

    Dim newVersion As String
    newVersion = InputBox("Informe a nova versao (ex.: 1.0.0):", APP_TITLE, IIf(Len(curVersion) = 0, "1.0.0", curVersion))
    newVersion = Trim$(newVersion)
    If Len(newVersion) = 0 Then Exit Sub

    Dim desc As String
    desc = InputBox("Descreva as alteracoes desta versao:", APP_TITLE, "Primeira versao")
    desc = Trim$(desc)

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Unprotect Password:=pwd

    Dim lr As ListRow
    Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 1).Value = newVersion
    lr.Range.Cells(1, 2).Value = Date
    lr.Range.Cells(1, 3).Value = Application.UserName
    lr.Range.Cells(1, 4).Value = desc

    ws.Range("B3").Value = newVersion
    ws.Columns.AutoFit
    ws.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub Version_RefreshCurrent()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = GetWs(SH_VERSOES)
    Dim lo As ListObject
    Set lo = ws.ListObjects(TB_VERSOES)
    ws.Range("B3").Value = Version_GetCurrent(lo)
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, APP_TITLE
End Sub

Private Function Version_GetCurrent(ByVal lo As ListObject) As String
    Version_GetCurrent = ""
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Version_GetCurrent = CStr(lo.DataBodyRange.Cells(lo.DataBodyRange.Rows.Count, 1).Value)
End Function

