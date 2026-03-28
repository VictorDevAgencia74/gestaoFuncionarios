Attribute VB_Name = "modConfigUI"
Option Explicit

Public Sub Config_Open()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetWs(SH_CONFIG)
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(CFG_RETRO_CODE_CELL).Select
    Exit Sub

ErrHandler:
    MsgBox "Nao foi possivel abrir a aba Config. Rode Setup_InitializeWorkbook e tente novamente.", vbExclamation, APP_TITLE
End Sub

Public Sub Config_Hide()
    On Error Resume Next
    GetWs(SH_CONFIG).Visible = xlSheetVeryHidden
End Sub

Public Sub Config_Show()
    On Error Resume Next
    GetWs(SH_CONFIG).Visible = xlSheetVisible
End Sub

Public Sub Config_GoToRetroAuthorization()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = GetWs(SH_CONFIG)
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(CFG_RETRO_ALLOW_DAYS_CELL).Select
    Exit Sub

ErrHandler:
    MsgBox "Nao foi possivel acessar a configuracao de autorizacao.", vbExclamation, APP_TITLE
End Sub

