Attribute VB_Name = "modTests"
Option Explicit

Public Sub Test_RunAll()
    On Error GoTo ErrHandler

    Setup_InitializeWorkbook
    Sample_GenerateData

    Dim wsA As Worksheet
    Set wsA = GetWs(SH_ALOC_FORM)
    wsA.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))

    wsA.Range("B3").Value = "F000001"
    wsA.Range("B4").Value = "R01"
    wsA.Range("B5").Value = Date
    wsA.Range("B6").Value = Date + 10
    wsA.Range("B7").Value = "Teste 1"
    wsA.Range("B9").Value = "NAO"
    wsA.Range("B10").ClearContents
    wsA.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True
    Allocation_SaveFromForm

    wsA.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsA.Range("B3").Value = "F000001"
    wsA.Range("B4").Value = "R02"
    wsA.Range("B5").Value = Date + 5
    wsA.Range("B6").Value = Date + 12
    wsA.Range("B7").Value = "Teste sobreposicao"
    wsA.Range("B9").Value = "NAO"
    wsA.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True
    Allocation_SaveFromForm

    Exit Sub

ErrHandler:
    MsgBox "Teste finalizado: " & Err.Description, vbExclamation, APP_TITLE
End Sub

