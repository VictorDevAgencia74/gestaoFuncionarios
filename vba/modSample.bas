Attribute VB_Name = "modSample"
Option Explicit

Public Sub Sample_GenerateData()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Setup_InitializeWorkbook

    Dim wsF As Worksheet
    Dim wsR As Worksheet
    Dim wsA As Worksheet
    Set wsF = GetWs(SH_FUNC_DB)
    Set wsR = GetWs(SH_REGIOES)
    Set wsA = GetWs(SH_ALOC_DB)

    Dim pwd As String
    pwd = CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    wsF.Unprotect Password:=pwd
    wsR.Unprotect Password:=pwd
    wsA.Unprotect Password:=pwd

    Dim loF As ListObject
    Dim loR As ListObject
    Dim loA As ListObject
    Set loF = wsF.ListObjects(TB_FUNC)
    Set loR = wsR.ListObjects(TB_REG)
    Set loA = wsA.ListObjects(TB_ALOC)

    If Not loF.DataBodyRange Is Nothing Then loF.DataBodyRange.Delete
    If Not loR.DataBodyRange Is Nothing Then loR.DataBodyRange.Delete
    If Not loA.DataBodyRange Is Nothing Then loA.DataBodyRange.Delete

    Dim i As Long
    For i = 1 To 10
        Dim lrR As ListRow
        Set lrR = loR.ListRows.Add
        lrR.Range.Cells(1, 1).Value = "R" & Format$(i, "00")
        lrR.Range.Cells(1, 2).Value = "Regiao " & i
        lrR.Range.Cells(1, 3).Value = "Endereco " & i & ", Cidade"
        lrR.Range.Cells(1, 4).Value = "Supervisor " & i
        lrR.Range.Cells(1, 5).Value = 6 + (i Mod 5)
    Next i

    Dim nomes1 As Variant
    Dim nomes2 As Variant
    nomes1 = Array("Ana", "Bruno", "Carla", "Diego", "Elisa", "Fabio", "Gisele", "Hugo", "Iris", "Joao", "Karen", "Lucas", "Marina", "Nicolas", "Otavio", "Paula", "Rafael", "Sonia", "Tiago", "Vanessa")
    nomes2 = Array("Silva", "Souza", "Oliveira", "Santos", "Lima", "Costa", "Pereira", "Carvalho", "Almeida", "Gomes", "Ribeiro", "Martins", "Araujo", "Barbosa", "Rocha")

    Randomize
    For i = 1 To 50
        Dim lrF As ListRow
        Set lrF = loF.ListRows.Add
        Dim empId As String
        empId = "F" & Format$(i, "000000")
        lrF.Range.Cells(1, 1).Value = empId
        lrF.Range.Cells(1, 2).Value = nomes1(Int(Rnd() * (UBound(nomes1) + 1))) & " " & nomes2(Int(Rnd() * (UBound(nomes2) + 1)))
        lrF.Range.Cells(1, 3).Value = Sample_GenerateCPF()
        lrF.Range.Cells(1, 4).Value = Date - Int(Rnd() * 900)
        lrF.Range.Cells(1, 5).Value = "Analista"
        lrF.Range.Cells(1, 6).Value = "Operacoes"
        If i Mod 13 = 0 Then
            lrF.Range.Cells(1, 7).Value = "Inativo"
        Else
            lrF.Range.Cells(1, 7).Value = "Ativo"
        End If
        lrF.Range.Cells(1, 8).Value = Now
        lrF.Range.Cells(1, 9).Value = Now
    Next i

    Dim empIdx As Long
    For empIdx = 1 To 40
        Dim regIdx As Long
        regIdx = 1 + (empIdx Mod 10)
        Dim dtIni As Date
        Dim dtFim As Date
        dtIni = Date - Int(Rnd() * 60)
        dtFim = dtIni + 30 + Int(Rnd() * 60)

        Dim lrA As ListRow
        Set lrA = loA.ListRows.Add
        lrA.Range.Cells(1, 1).Value = "A-" & NewGuidId()
        lrA.Range.Cells(1, 2).Value = "F" & Format$(empIdx, "000000")
        lrA.Range.Cells(1, 3).Value = "R" & Format$(regIdx, "00")
        lrA.Range.Cells(1, 4).Value = dtIni
        lrA.Range.Cells(1, 5).Value = dtFim
        lrA.Range.Cells(1, 6).Value = "Gerado automaticamente"
        lrA.Range.Cells(1, 7).Value = Now
        lrA.Range.Cells(1, 8).Value = Application.UserName
    Next empIdx

    wsF.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    wsR.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True
    wsA.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True

    Setup_RefreshAfterDataChange
    Application.ScreenUpdating = True
    MsgBox "Dados simulados criados (50 funcionarios, 10 regioes).", vbInformation, APP_TITLE
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, APP_TITLE
End Sub

Private Function Sample_GenerateCPF() As String
    Dim base(1 To 9) As Integer
    Dim i As Long
    For i = 1 To 9
        base(i) = Int(Rnd() * 10)
    Next i

    Dim sum As Long
    sum = 0
    For i = 1 To 9
        sum = sum + base(i) * (11 - i)
    Next i
    Dim d1 As Integer
    d1 = (sum * 10) Mod 11
    If d1 = 10 Then d1 = 0

    sum = 0
    For i = 1 To 9
        sum = sum + base(i) * (12 - i)
    Next i
    sum = sum + d1 * 2
    Dim d2 As Integer
    d2 = (sum * 10) Mod 11
    If d2 = 10 Then d2 = 0

    Dim out As String
    out = ""
    For i = 1 To 9
        out = out & CStr(base(i))
    Next i
    out = out & CStr(d1) & CStr(d2)
    Sample_GenerateCPF = out
End Function

