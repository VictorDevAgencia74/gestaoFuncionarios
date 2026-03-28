Attribute VB_Name = "modQuery"
Option Explicit

Public Sub Query_Clear()
    Dim ws As Worksheet
    Set ws = GetWs(SH_CONSULTA)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    ws.Range("B3:B6").ClearContents

    Dim lo As ListObject
    Set lo = ws.ListObjects(TB_QUERY)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete

    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Public Sub Query_Run()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = GetWs(SH_CONSULTA)
    Dim filtroFunc As String
    Dim filtroReg As String
    Dim dtIniV As Variant
    Dim dtFimV As Variant
    filtroFunc = Trim$(CStr(ws.Range("B3").Value))
    filtroReg = UCase$(Trim$(CStr(ws.Range("B4").Value)))
    dtIniV = ws.Range("B5").Value
    dtFimV = ws.Range("B6").Value

    Dim hasDtIni As Boolean
    Dim hasDtFim As Boolean
    hasDtIni = IsDate(dtIniV)
    hasDtFim = IsDate(dtFimV)

    Dim dtIni As Date
    Dim dtFim As Date
    If hasDtIni Then dtIni = CDate(dtIniV)
    If hasDtFim Then dtFim = CDate(dtFimV)
    If hasDtIni And hasDtFim Then
        If dtIni > dtFim Then Err.Raise vbObjectError + 400, APP_TITLE, "Periodo invalido na consulta."
    End If

    Dim loOut As ListObject
    Set loOut = ws.ListObjects(TB_QUERY)
    ws.Unprotect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL))
    If Not loOut.DataBodyRange Is Nothing Then loOut.DataBodyRange.Delete

    Dim loA As ListObject
    Set loA = GetWs(SH_ALOC_DB).ListObjects(TB_ALOC)
    If loA.DataBodyRange Is Nothing Then
        ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
        Exit Sub
    End If

    Dim r As Long
    Dim idxAid As Long, idxEmp As Long, idxReg As Long, idxIni As Long, idxFim As Long, idxObs As Long
    idxAid = TableColIndex(loA, "AlocacaoID")
    idxEmp = TableColIndex(loA, "FuncionarioID")
    idxReg = TableColIndex(loA, "RegiaoCodigo")
    idxIni = TableColIndex(loA, "DataInicio")
    idxFim = TableColIndex(loA, "DataFim")
    idxObs = TableColIndex(loA, "Observacoes")

    For r = 1 To loA.DataBodyRange.Rows.Count
        Dim empId As String
        Dim regCode As String
        Dim di As Date
        Dim df As Date
        empId = CStr(loA.DataBodyRange.Cells(r, idxEmp).Value)
        regCode = CStr(loA.DataBodyRange.Cells(r, idxReg).Value)
        di = CDate(loA.DataBodyRange.Cells(r, idxIni).Value)
        df = CDate(loA.DataBodyRange.Cells(r, idxFim).Value)

        If Len(filtroReg) > 0 Then
            If StrComp(regCode, filtroReg, vbTextCompare) <> 0 Then GoTo NextR
        End If

        If hasDtIni Then
            If df < dtIni Then GoTo NextR
        End If
        If hasDtFim Then
            If di > dtFim Then GoTo NextR
        End If

        If Len(filtroFunc) > 0 Then
            Dim nm As String
            nm = Employee_GetName(empId)
            If (StrComp(empId, filtroFunc, vbTextCompare) <> 0) And (InStr(1, nm, filtroFunc, vbTextCompare) = 0) Then GoTo NextR
        End If

        Dim outRow As ListRow
        Set outRow = loOut.ListRows.Add
        With outRow.Range
            .Cells(1, 1).Value = CStr(loA.DataBodyRange.Cells(r, idxAid).Value)
            .Cells(1, 2).Value = empId
            .Cells(1, 3).Value = Employee_GetName(empId)
            .Cells(1, 4).Value = Employee_GetCPF(empId)
            .Cells(1, 5).Value = regCode
            .Cells(1, 6).Value = Region_GetName(regCode)
            .Cells(1, 7).Value = di
            .Cells(1, 8).Value = df
            .Cells(1, 9).Value = CStr(loA.DataBodyRange.Cells(r, idxObs).Value)
        End With

NextR:
    Next r

    ws.Columns.AutoFit
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub
ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    ws.Protect Password:=CStr(GetConfigValue(CFG_PROTECT_PWD_CELL)), UserInterfaceOnly:=True, AllowFiltering:=True
    Dim msg As String
    msg = errDesc
    If Len(msg) = 0 Then msg = "Erro " & CStr(errNum)
    MsgBox msg, vbExclamation, APP_TITLE
End Sub

