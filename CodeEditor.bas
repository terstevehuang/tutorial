Attribute VB_Name = "CodeEditor"
Option Explicit

Public Sub AppendString1()

    Dim rng         As Range
    Dim strArray()  As String
    Dim i           As Integer
    Dim TempStr     As String
    
    For Each rng In Selection
        TempStr = ""
        strArray = Split(rng.Value, "_")

        For i = 0 To UBound(strArray) - 4
            TempStr = TempStr & strArray(i) & "_"
        Next i
        
        rng.offset(0, 2).Value = Mid(TempStr, 1, Len(TempStr) - 1)
        rng.offset(0, 3).Value = strArray(UBound(strArray) - 3)
        rng.offset(0, 4).Value = strArray(UBound(strArray) - 1) & "_" & strArray(UBound(strArray))
    Next rng
    
End Sub


Public Sub ss()
'update ss function
    Dim rng As Range
    For Each rng In Selection
        rng.Value = Replace(rng.Value, "Gpu", "Gfx")
    Next rng
End Sub


Public Sub overKillCnt()
    
    Dim rng As Range
    Dim overKillCnt     As Integer:
    Dim gradeChangeCnt  As Integer:
    
    Dim colMax          As Integer: colMax = Range("A1").End(xlToRight).Column
    Dim rowMax          As Integer: rowMax = Range("A1").End(xlDown).row
    Dim colIter         As Integer
    Dim rowIter         As Integer
    Dim rowCnt          As Integer

'
'    For rowIter = 2 To rowMax
'        overKillCnt = 0
'        gradeChangeCnt = 0
'        If InStr(Range("A" & rowIter).Value, "_50MHz") Then
'            For Each rng In Selection
'                rowCnt = rng.Row
'                If rng.Value <> rng.Offset(-1, 0).Value Then gradeChangeCnt = gradeChangeCnt + 1
'                If rng.Value = "" Then overKillCnt = overKillCnt + 1
'            Next rng
'        End If
'
'    Next rowIter
    
'    For Each rng In Selection
'        rowCnt = rng.Row
'        If rng.Value <> rng.Offset(-1, 0).Value Then gradeChangeCnt = gradeChangeCnt + 1
'        If rng.Value = "" Then overKillCnt = overKillCnt + 1
'    Next rng
    
    For Each rng In Selection
        rowCnt = rng.row
        If rng.Value <> rng.offset(-2, 0).Value Then gradeChangeCnt = gradeChangeCnt + 1
        If rng.Value = 0 Then overKillCnt = overKillCnt + 1
    Next rng
'
    Range("B" & rowCnt).Value = overKillCnt
    Range("C" & rowCnt).Value = gradeChangeCnt
    

End Sub

Public Sub AppendString2()

    Dim rng As Range
    Dim strArray() As String
    Dim i As Integer
    
    For Each rng In Selection
    
        strArray = Split(rng.Value, "_")
        rng.Value = ""
        For i = 0 To UBound(strArray) - 1
            rng.Value = rng.Value & strArray(i) & "_"
        Next i
        rng.Value = Mid(rng.Value, 1, Len(rng.Value) - 1) 'rng.Value & "Diag_" & strArray(UBound(strArray))
        
    Next rng
    
End Sub

Public Sub spp()

    Dim rng As Range
    Dim strArray() As String
    For Each rng In Selection
        strArray = Split(rng.Value, " ")
        rng.Value = strArray(0)
        rng.offset(0, 1).Value = strArray(1)
    Next rng

End Sub


Public Sub AppendString()

    Dim rng As Range
    
    For Each rng In Selection
        rng.Value = rng.Value & "_UVS_NV"
    Next rng

    
End Sub


Public Sub ReplaceString()

    Dim rng As Range
    For Each rng In Selection
        rng.Value = Replace(rng.Value, "_Shm", "_Diag_Shm")
    Next rng

End Sub



Public Sub ReplaceString00()

    Dim rng As Range
    For Each rng In Selection
        rng.Value = Replace(rng.Value, "_HV", "_Diag_HV")
    Next rng

End Sub

Public Sub ReplaceString123()

    Dim rng As Range
    For Each rng In Selection
        rng.Value = Replace(rng.Value, "SocTD", "SocSA")
    Next rng

End Sub

Public Sub InsertString()

    Dim rng As Range
    For Each rng In Selection
        rng.Value = "F_soctd_p1_LV," & rng.Value
    Next rng

End Sub

Public Sub PatrialString()

    Dim rng As Range
    Dim strArray() As String
    Dim i As Integer
    
    For Each rng In Selection
        strArray = Split(rng.Value, ",")
        rng.Value = strArray(1)
    Next rng
    
End Sub

Public Sub FindInstList()

    Dim instList    As String
    Dim sh          As Worksheet
    
    For Each sh In Worksheets
        If InStr(LCase(sh.Name), "inst") Then
            instList = instList + sh.Name + ","
        End If
    Next sh
    
    Debug.Print instList
    
End Sub

Public Sub FindPSetList()

    Dim pSetList    As String
    Dim sh          As Worksheet
    
    For Each sh In Worksheets
        If InStr(LCase(sh.Name), "pat") Then
            pSetList = pSetList + sh.Name + ","
        End If
    Next sh
    
    Debug.Print instList
    
End Sub




