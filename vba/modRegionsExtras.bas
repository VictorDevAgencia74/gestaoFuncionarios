Attribute VB_Name = "modRegionsExtras"
Option Explicit

Public Function Region_GetSupervisor(ByVal codigo As String) As String
    Dim lo As ListObject
    Set lo = GetWs(SH_REGIOES).ListObjects(TB_REG)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim idxCode As Long
    Dim idxSup As Long
    idxCode = TableColIndex(lo, "RegiaoCodigo")
    idxSup = TableColIndex(lo, "Supervisor")
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, idxCode).Value), codigo, vbTextCompare) = 0 Then
            Region_GetSupervisor = CStr(lo.DataBodyRange.Cells(r, idxSup).Value)
            Exit Function
        End If
    Next r
End Function

