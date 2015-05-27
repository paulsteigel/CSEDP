Option Explicit

Sub GetObjSource(ObjControl As Control, Optional ParrentID As String = "", _
    Optional colCount As Long = 2, Optional RowSourceName As String = "", _
    Optional SearchCell As String = "", Optional ResourceText As String)
    'Fill in Commbo or listbox with region table
    On Error GoTo err_handler
    Dim Arr() As Variant
    If RowSourceName <> "" Then
        ' This will die when there is only one cell...
        If Range(RowSourceName).Cells.Count = 1 Then
            Dim tmpArr(1, 1)
            tmpArr(1, 1) = Range(RowSourceName)
            Arr = tmpArr
        Else
            Arr = Range(RowSourceName)
        End If
    Else
        Arr = Range("tblRegions")
    End If
    Dim R As Long
    With ObjControl
        .ColumnCount = colCount
        .ColumnWidths = IIf(colCount = 1, .Width - 10, "0;" & .Width - 10)
        .Clear
        
        For R = 1 To UBound(Arr, 1) ' First array dimension is rows.
            If ParrentID = "" And RowSourceName <> "" Then
                If Arr(R, 1) <> "" And Not Arr(R, 1) Like "<<*" Then
                    .AddItem Arr(R, 1)
                    ResourceText = ResourceText & "[" & Arr(R, 1) & "]"
                    If colCount = 2 Then
                        .List(.ListCount - 1, 1) = Arr(R, 2)
                    End If
                End If
                If Arr(R, 1) = SearchCell Then
                    If Not Arr(1, 1) Like "<<*" Then
                        .Selected(R - 1) = True
                    Else
                        .Selected(R - 2) = True
                    End If
                End If
            Else
                If Arr(R, 3) = ParrentID Then
                    If colCount = 2 Then
                        .AddItem Arr(R, 1)
                        .List(.ListCount - 1, 1) = Arr(R, 4)
                    Else
                        .AddItem Arr(R, 4)
                    End If
                End If
            End If
        Next R
    End With
err_handler:
End Sub

Function GetAbrFromText(TextString As String) As String
    ' To get just first letter of the text string
    Dim i As Long, rStr As String
    TextString = Trim(TextString)
    rStr = Left(Trim(TextString), 1)
    i = InStr(TextString, " ")
    If i <= 0 Then
        rStr = rStr & "BL"
        GoTo ExitFunc
    End If
    While i > 0
        rStr = rStr & Mid(TextString, i + 1, 1)
        i = InStr(i + 1, TextString, " ")
    Wend
ExitFunc:
    GetAbrFromText = rStr
End Function

Sub TryExt()
    'Fill in Commbo or listbox with region table
    Dim theCell As Range, i As Long
    ShowOff
    For i = 1 To Range("tblRegions").Rows.Count
        Range("tblRegions").Cells(i, 4) = Trim(Range("tblRegions").Cells(i, 4))
    Next
    ShowOff True
End Sub
