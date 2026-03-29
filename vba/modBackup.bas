Attribute VB_Name = "modBackup"
Option Explicit

Public Sub Backup_CreateNow()
    Backup_Create True
End Sub

Public Sub Backup_HandleBeforeClose(ByRef Cancel As Boolean)
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Deseja fazer um backup antes de fechar?", vbQuestion + vbYesNoCancel, APP_TITLE)
    If resp = vbCancel Then
        Cancel = True
        Exit Sub
    End If
    If resp = vbYes Then
        On Error GoTo ErrHandler
        Backup_Create False
    End If
    Exit Sub
ErrHandler:
    MsgBox "Falha ao criar backup: " & Err.Description, vbExclamation, APP_TITLE
    Cancel = True
End Sub

Public Sub Backup_Create(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim folderPath As String
    folderPath = Backup_FolderPath()
    Dim filePath As String
    filePath = folderPath & "\" & Backup_FileName()
    ThisWorkbook.SaveCopyAs filePath
    If showMessage Then MsgBox "Backup criado em:" & vbCrLf & filePath, vbInformation, APP_TITLE
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, APP_TITLE, Err.Description
End Sub

Public Sub Backup_Import()
    On Error GoTo ErrHandler
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Importar um backup vai substituir os dados atuais (Funcionarios, Regioes, Alocacoes e Config). Continuar?", vbExclamation + vbYesNo, APP_TITLE)
    If resp <> vbYes Then Exit Sub

    Dim fd As Object
    Const FILE_PICKER As Long = 3
    Set fd = Application.FileDialog(FILE_PICKER)
    fd.AllowMultiSelect = False
    fd.Title = "Selecione um arquivo de backup"
    fd.InitialFileName = Backup_FolderPath() & "\"
    On Error Resume Next
    fd.Filters.Clear
    fd.Filters.Add "Backups Excel", "*.xlsm;*.xltm;*.xlsb;*.xlsx"
    fd.Filters.Add "Todos os arquivos", "*.*"
    On Error GoTo ErrHandler

    If fd.Show <> -1 Then Exit Sub
    Dim filePath As String
    filePath = CStr(fd.SelectedItems(1))
    Backup_ImportFromFile filePath
    Exit Sub
ErrHandler:
    MsgBox "Nao foi possivel importar o backup: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub Backup_ImportFromFile(ByVal filePath As String)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Backup_Create False

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)

    Dim pwd As String
    Dim srcPwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    srcPwd = Backup_TryGetWorkbookConfigPassword(srcWb)
    Backup_UnprotectAll Array(pwd, srcPwd, "alocacao", vbNullString)
    Backup_ValidateDestinationReady

    Backup_CopyAllDataSafely srcWb
    srcWb.Close SaveChanges:=False

    Setup_RefreshAfterDataChange

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Backup importado com sucesso." & vbCrLf & Backup_ImportSummary(), vbInformation, APP_TITLE
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, APP_TITLE, Err.Description
End Sub

Private Sub Backup_CopyAllDataSafely(ByVal srcWb As Workbook)
    Dim funcData As Variant
    Dim regData As Variant
    Dim alocData As Variant
    Dim configPwd As Variant
    Dim configRetroCode As Variant
    Dim configRetroDays As Variant
    Dim configWarnDays As Variant
    Dim colA As Variant
    Dim colD As Variant
    Dim colG As Variant

    funcData = Backup_ReadTableData(srcWb.Worksheets(SH_FUNC_DB).ListObjects(TB_FUNC))
    regData = Backup_ReadTableData(srcWb.Worksheets(SH_REGIOES).ListObjects(TB_REG))
    alocData = Backup_ReadTableData(srcWb.Worksheets(SH_ALOC_DB).ListObjects(TB_ALOC))

    With srcWb.Worksheets(SH_CONFIG)
        configPwd = .Range(CFG_PROTECT_PWD_CELL).Value
        configRetroCode = .Range(CFG_RETRO_CODE_CELL).Value
        configRetroDays = .Range(CFG_RETRO_ALLOW_DAYS_CELL).Value
        configWarnDays = .Range(CFG_EXPIRY_WARN_DAYS_CELL).Value
        colA = .Range(.Cells(6, 1), .Cells(100, 1)).Value
        colD = .Range(.Cells(6, 4), .Cells(100, 4)).Value
        colG = .Range(.Cells(6, 7), .Cells(100, 7)).Value
    End With

    Backup_WriteTableData GetWs(SH_FUNC_DB).ListObjects(TB_FUNC), funcData
    Backup_WriteTableData GetWs(SH_REGIOES).ListObjects(TB_REG), regData
    Backup_WriteTableData GetWs(SH_ALOC_DB).ListObjects(TB_ALOC), alocData
    Backup_WriteConfig configPwd, configRetroCode, configRetroDays, configWarnDays, colA, colD, colG
End Sub

Private Function Backup_ReadTableData(ByVal srcLo As ListObject) As Variant
    If srcLo.DataBodyRange Is Nothing Then
        Backup_ReadTableData = Empty
    Else
        Backup_ReadTableData = srcLo.DataBodyRange.Value
    End If
End Sub

Private Sub Backup_WriteTableData(ByVal destLo As ListObject, ByVal data As Variant)
    On Error GoTo ErrHandler
    Dim targetRows As Long
    Dim targetCols As Long

    If IsEmpty(data) Then
        targetRows = 0
    ElseIf IsArray(data) Then
        targetRows = UBound(data, 1)
        targetCols = UBound(data, 2)
    Else
        targetRows = 1
        targetCols = 1
    End If

    If Not destLo.DataBodyRange Is Nothing Then
        destLo.DataBodyRange.Delete
    End If

    If targetRows = 0 Then Exit Sub

    Dim i As Long
    For i = 1 To targetRows
        destLo.ListRows.Add
    Next i

    destLo.DataBodyRange.Resize(targetRows, targetCols).Value = data
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, APP_TITLE, Err.Description
End Sub

Private Sub Backup_WriteConfig(ByVal protectPwd As Variant, ByVal retroCode As Variant, ByVal retroDays As Variant, ByVal warnDays As Variant, ByVal colA As Variant, ByVal colD As Variant, ByVal colG As Variant)
    Dim destWs As Worksheet
    Set destWs = GetWs(SH_CONFIG)

    destWs.Range(CFG_PROTECT_PWD_CELL).Value = protectPwd
    destWs.Range(CFG_RETRO_CODE_CELL).Value = retroCode
    destWs.Range(CFG_RETRO_ALLOW_DAYS_CELL).Value = retroDays
    destWs.Range(CFG_EXPIRY_WARN_DAYS_CELL).Value = warnDays
    destWs.Range(destWs.Cells(6, 1), destWs.Cells(100, 1)).Value = colA
    destWs.Range(destWs.Cells(6, 4), destWs.Cells(100, 4)).Value = colD
    destWs.Range(destWs.Cells(6, 7), destWs.Cells(100, 7)).Value = colG
End Sub

Private Function Backup_FolderPath() As String
    Backup_FolderPath = EnsureFolder(WorkbookFolder() & "\" & "bkp")
End Function

Private Function Backup_FileName() As String
    Dim baseName As String
    Dim ext As String
    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)

    ext = "xlsm"
    If InStrRev(ThisWorkbook.Name, ".") > 0 Then
        ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") + 1)
    End If
    If StrComp(LCase$(ext), "xltm", vbTextCompare) = 0 Then ext = "xlsm"
    If StrComp(LCase$(ext), "xltx", vbTextCompare) = 0 Then ext = "xlsx"

    Backup_FileName = baseName & "_bkp_" & Format$(Now, "yyyy-mm-dd_hhnnss") & "." & ext
End Function

Private Sub Backup_ValidateDestinationReady()
    If Not WorksheetExists(SH_FUNC_DB) Then Err.Raise vbObjectError + 901, APP_TITLE, "Aba '" & SH_FUNC_DB & "' nao existe. Rode Setup_InitializeWorkbook."
    If Not WorksheetExists(SH_ALOC_DB) Then Err.Raise vbObjectError + 902, APP_TITLE, "Aba '" & SH_ALOC_DB & "' nao existe. Rode Setup_InitializeWorkbook."
    If Not WorksheetExists(SH_REGIOES) Then Err.Raise vbObjectError + 903, APP_TITLE, "Aba '" & SH_REGIOES & "' nao existe. Rode Setup_InitializeWorkbook."
    If Not WorksheetExists(SH_CONFIG) Then Err.Raise vbObjectError + 904, APP_TITLE, "Aba '" & SH_CONFIG & "' nao existe. Rode Setup_InitializeWorkbook."

    If Not TableExists(GetWs(SH_FUNC_DB), TB_FUNC) Then Err.Raise vbObjectError + 905, APP_TITLE, "Tabela '" & TB_FUNC & "' nao existe. Rode Setup_InitializeWorkbook."
    If Not TableExists(GetWs(SH_ALOC_DB), TB_ALOC) Then Err.Raise vbObjectError + 906, APP_TITLE, "Tabela '" & TB_ALOC & "' nao existe. Rode Setup_InitializeWorkbook."
    If Not TableExists(GetWs(SH_REGIOES), TB_REG) Then Err.Raise vbObjectError + 907, APP_TITLE, "Tabela '" & TB_REG & "' nao existe. Rode Setup_InitializeWorkbook."
End Sub

Private Sub Backup_UnprotectAll(ByVal candidatePasswords As Variant)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Backup_UnprotectSheet ws, candidatePasswords
    Next ws
End Sub

Private Sub Backup_UnprotectSheet(ByVal ws As Worksheet, ByVal candidatePasswords As Variant)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    Dim i As Long
    For i = LBound(candidatePasswords) To UBound(candidatePasswords)
        On Error Resume Next
        ws.Unprotect Password:=CStr(candidatePasswords(i))
        If Err.Number = 0 Then Exit For
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

Private Function Backup_TryGetWorkbookConfigPassword(ByVal wb As Workbook) As String
    On Error GoTo ErrHandler
    Backup_TryGetWorkbookConfigPassword = CStr(wb.Worksheets(SH_CONFIG).Range(CFG_PROTECT_PWD_CELL).Value)
    Exit Function
ErrHandler:
    Backup_TryGetWorkbookConfigPassword = vbNullString
End Function

Private Function Backup_TableRowCount(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then
        Backup_TableRowCount = 0
    Else
        Backup_TableRowCount = lo.DataBodyRange.Rows.Count
    End If
End Function

Private Function Backup_ImportSummary() As String
    On Error GoTo ErrHandler
    Dim nFunc As Long
    Dim nReg As Long
    Dim nAlo As Long

    nFunc = Backup_TableRowCount(GetWs(SH_FUNC_DB).ListObjects(TB_FUNC))
    nReg = Backup_TableRowCount(GetWs(SH_REGIOES).ListObjects(TB_REG))
    nAlo = Backup_TableRowCount(GetWs(SH_ALOC_DB).ListObjects(TB_ALOC))

    Backup_ImportSummary = "Registros no arquivo atual:" & vbCrLf & _
        " - Funcionarios: " & CStr(nFunc) & vbCrLf & _
        " - Regioes: " & CStr(nReg) & vbCrLf & _
        " - Alocacoes: " & CStr(nAlo)
    Exit Function
ErrHandler:
    Backup_ImportSummary = vbNullString
End Function
