Option Explicit

Sub UpdateChange(ByVal Target As Range, ByVal WrkSheet As Worksheet, Optional ByVal CmdObject As Object = Nothing)
    ' this process is to make the update buton...
    'If isFrmLoaded Then Exit Sub
    ShowOff
    With Target
        ' Check if a shape has been created or not?
        If (.Row > 6 And .Row <= 555) And (.Column = 12 Or .Column = 14) Then
            CmdObject.Top = ActiveCell.Top + (ActiveCell.Height - CmdObject.Height)
            CmdObject.Left = ActiveCell.Left - CmdObject.Width
            CmdObject.Visible = msoTrue
        Else
            CmdObject.Visible = msoFalse
        End If
    End With
    ShowOff True
End Sub
