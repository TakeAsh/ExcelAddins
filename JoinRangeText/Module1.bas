Attribute VB_Name = "Module1"

Sub RegisterJoinRangeText()
    Application.MacroOptions _
        Macro:="JoinRangeText", _
        Description:="Join texts in selected range", _
        Category:=7, _
        ArgumentDescriptions:=Array( _
            "Range to be joined", _
            "Delimiter (optional)" _
        )
End Sub

Public Function JoinRangeText(rng As Range, Optional delim As String) As String
Attribute JoinRangeText.VB_Description = "Join texts in selected range"
Attribute JoinRangeText.VB_ProcData.VB_Invoke_Func = " \n7"
    Dim work As String
    Dim c As Range
    work = ""
    For Each c In rng
        work = work & c.Text & delim
    Next
    JoinRangeText = Left(work, Len(work) - Len(delim))
End Function

Sub RegisterJoinRangeTextA()
    Application.MacroOptions _
        Macro:="JoinRangeTextA", _
        Description:="Join not empty texts in selected range", _
        Category:=7, _
        ArgumentDescriptions:=Array( _
            "Range to be joined", _
            "Delimiter (optional)" _
        )
End Sub

Public Function JoinRangeTextA(rng As Range, Optional delim As String) As String
Attribute JoinRangeTextA.VB_Description = "Join not empty texts in selected range"
Attribute JoinRangeTextA.VB_ProcData.VB_Invoke_Func = " \n7"
    Dim work As String
    Dim c As Range
    work = ""
    For Each c In rng
        If c.Text <> "" Then
            work = work & c.Text & delim
        End If
    Next
    If Len(work) > 0 Then
        work = Left(work, Len(work) - Len(delim))
    End If
    JoinRangeTextA = work
End Function

Sub RegisterJoinRangeValue()
    Application.MacroOptions _
        Macro:="JoinRangeValue", _
        Description:="Join values in selected range", _
        Category:=7, _
        ArgumentDescriptions:=Array( _
            "Range to be joined", _
            "Delimiter (optional)" _
        )
End Sub

Public Function JoinRangeValue(rng As Range, Optional delim As String) As String
Attribute JoinRangeValue.VB_Description = "Join values in selected range"
Attribute JoinRangeValue.VB_ProcData.VB_Invoke_Func = " \n7"
    Dim work As String
    Dim c As Range
    work = ""
    For Each c In rng
        work = work & c.Value & delim
    Next
    JoinRangeValue = Left(work, Len(work) - Len(delim))
End Function

Sub RegisterJoinRangeValueA()
    Application.MacroOptions _
        Macro:="JoinRangeValueA", _
        Description:="Join not empty values in selected range", _
        Category:=7, _
        ArgumentDescriptions:=Array( _
            "Range to be joined", _
            "Delimiter (optional)" _
        )
End Sub

Public Function JoinRangeValueA(rng As Range, Optional delim As String) As String
Attribute JoinRangeValueA.VB_Description = "Join not empty values in selected range"
Attribute JoinRangeValueA.VB_ProcData.VB_Invoke_Func = " \n7"
    Dim work As String
    Dim c As Range
    work = ""
    For Each c In rng
        If c.Value <> "" Then
            work = work & c.Value & delim
        End If
    Next
    If Len(work) > 0 Then
        work = Left(work, Len(work) - Len(delim))
    End If
    JoinRangeValueA = work
End Function
