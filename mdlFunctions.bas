Option Explicit
 
Function FalseInput(CtrlName As Control) As Boolean
    Dim tData As String
    If CtrlName = "" Then Exit Function
    If Not IsDate(CtrlName) Then GoTo tCont
    tData = InputDate(CtrlName)
    If Not tData Like "12:00*" Then Exit Function
tCont:
    CtrlName = ""
    CtrlName.SetFocus
    FalseInput = True
End Function

Function InputDate(iDateStr As Variant) As Date
    ' Send data piece from database to console
    ' default the data will from db to console, output shall be formated
    ' input shall be converted back to serial date
    Dim iStr As String, iSpliter As Variant
    
    On Error GoTo ErrHandler
    iSpliter = Split(iDateStr, "/")
    If UBound(iSpliter) < 2 Then GoTo ErrHandler
    ' Now we have to see what locale we are now at
    InputDate = DateSerial(iSpliter(2), iSpliter(0), iSpliter(1))
ErrHandler:
End Function

'===================================================
' For form level object
Private Sub ArchiveConfig()
    ExternalLoad = True
    Dim MyCtl As Control, i As Long, j As Long
    Dim MyForm As UserForm
    Dim MyCell As Range, tmpCell As Range
    Set MyCell = Range("tblFormConfig").Offset(1)
    Set MyForm = frmMain
    
    '1. Search for area to keep data
    For Each MyCtl In MyForm.Controls
        MyCell = "FORM_frmMain"
        MyCell.Offset(, 2) = MyCtl.Name
        'If TypeOf MyCtl Is MultiPage Then
        '    MyCell.Offset(, 1) = 1
        '    Set tmpCell = MyCell
        '    For j = 0 To MyCtl.Pages.Count - 1
        '        tmpCell.Offset(, j + 3) = MyCtl.Pages(j).Caption
        '    Next
        'Else
        '    MyCell.Offset(, 1) = 0
        '    MyCell.Offset(, 3) = GetCaption(MyCtl)
        'End If
        Set MyCell = MyCell.Offset(1)
    Next
End Sub

Property Get SetConfig(ObjName As String, FrmObj As UserForm) As String
    Dim j As Long
    Dim MyCell As Range, tmpCell As Range
    Set MyCell = Range("tblFormConfig").Offset(1)
    '1. Search for area to keep data
    While MyCell <> "" And MyCell <> ObjName
        Set MyCell = MyCell.Offset(1)
    Wend
    While MyCell = ObjName
        Select Case MyCell.Offset(, 1)
        Case 0:
            SetCaption FrmObj.Controls(CStr(MyCell.Offset(, 2))), MyCell.Offset(, 3), MyCell.Offset(, 4)
        Case 2, 4: ' Just set tag value
            FrmObj.Controls(CStr(MyCell.Offset(, 2))).Tag = MyCell.Offset(, 4)
            SetCaption FrmObj.Controls(CStr(MyCell.Offset(, 2))), MyCell.Offset(, 3), MyCell.Offset(, 4)
        Case 3:
            ' for form caption
            SetConfig = MyCell.Offset(, 3)
        Case Else
            Set tmpCell = MyCell
            For j = 0 To FrmObj.Controls(CStr(MyCell.Offset(, 2))).Pages.Count - 1
                FrmObj.Controls(CStr(MyCell.Offset(, 2))).Pages(j).Caption = tmpCell.Offset(, j + 3)
            Next
        End Select
        Set MyCell = MyCell.Offset(1)
    Wend
    Set MyCell = Nothing
    Set tmpCell = Nothing
End Property

Private Sub SetCaption(MyObj As Object, iCaption As String, Optional ControlTipStr As String = "")
    If iCaption <> "" Then MyObj.Caption = iCaption
    If ControlTipStr <> "" Then MyObj.ControlTipText = ControlTipStr
End Sub

Private Function GetCaption(Obj As Object) As String
    On Error GoTo ErrHandler
    GetCaption = Obj.Caption
ErrHandler:
End Function

Sub ToggleFilterKey()
    ' This shall help to disable filter
    If Not ActiveSheet.FilterMode Then
        QuickFilter
    Else
        ShowAll ActiveSheet
        ' Repair sheet if neccessary
        RepairSheet ActiveSheet.Name
    End If
End Sub

Sub InsertVillage()
    If ActiveSheet.Name <> "II.2.A" Then Exit Sub
    If MsgBox(MSG("MSG_ADD_VILLAGE"), vbQuestion + vbYesNo) = vbYes Then
        Dim TheRange As Range
        Set TheRange = AddRevVillage(1)
        ShowOff
        ModifyColumns
        ShowOff True
        ' Get to Data table for putting village name
        Sheets("Data").Activate
        TheRange.Activate
    End If
End Sub

Sub RemoveVillage()
    If ActiveSheet.Name <> "II.2.A" Then Exit Sub
    ' if just remain 2 colums - dont allow removal
    If Range("RNG_II2A").Column - Range("RNG_IIAST").Column = 6 Then
        MsgBox MSG("MSG_REMOVE_VILLAGE_DISALLOW"), vbCritical
        Exit Sub
    End If
    If MsgBox(Replace(MSG("MSG_REMOVE_VILLAGE"), "%s%", Sheet4.Range("RNG_II2A").Offset(0, -1)), vbQuestion + vbYesNo) = vbYes Then
        Call AddRevVillage(-1)
        ShowOff
        ModifyColumns -1
        ShowOff True
    End If
End Sub

Private Function AddRevVillage(param As Long) As Range
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Data").Range("tblVillageStart")
    While Len(Trim(rng)) > 0
        Set rng = rng.Offset(1)
    Wend
    ' Now I am at the last point
    If param < 0 Then
        rng.Offset(-2) = rng.Offset(-1)
        rng.Offset(-1) = ""
    Else
        rng = MSG("MSG_VIL_NEW")
        Set AddRevVillage = rng
    End If
    Set rng = Nothing
End Function

Sub ShowAll(SheetObj As Worksheet)
    On Error Resume Next
    XUnProtectSheet SheetObj
    SheetObj.ShowAllData
    XProtectSheet SheetObj
    SheetObj.Range("A8").Activate
End Sub

Sub ShowSelectForm()
    ' This shall display a form for selecting something
    Dim isSelected As Boolean
    Select Case ActiveSheet.Name
    Case "II.5.B"
        With ActiveCell
            If .Row > 6 And .Row <= 555 Then
                Select Case .Column
                Case 8:    ' Location for activity
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_ADD_LOCATION_VILLAGE")
                        .DataSetName = MSG("MSG_SELECT_VILLAGE")
                        .DataSource = "tblVillage"
                        .ReadOnly = True
                        .WrapOutput = True
                    End With
                Case 15:    ' Funding source
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_BUDGET")
                        .DataSource = "Nguonvon"
                        .ModifyColumn = True
                    End With
                Case 17:    ' Unit in charge
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_UNIT_INCHARGE")
                        .DataSource = "Bannganh"
                    End With
                Case 18:    ' Category
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_SELECT_ACT_CATEGORY")
                        .DataSource = "LST_CATEGORY|"
                        '.SpecialNote = MSG("VAL_STATUS")
                        '.ReadOnly = True
                    End With
                Case 19:    ' Status
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_SELECT_STATUS_ACT")
                        .DataSource = "Dexuat"
                        .SpecialNote = MSG("VAL_STATUS")
                        .ReadOnly = True
                    End With
                Case Else
                    isSelected = True
                End Select
                If Not isSelected Then frmSelect.Show vbModal
            End If
        End With
    Case "II.5.A":
        With ActiveCell
            If .Row > 5 And .Row <= 386 Then
                Select Case .Column
                Case 1:    ' Sector
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_SECTOR")
                        .DataSource = "Linhvuc"
                        .ReadOnly = True
                    End With
                Case 2:    ' Unit in charge
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_UNIT_INCHARGE")
                        .DataSource = "Bannganh"
                    End With
                Case Else
                    isSelected = True
                End Select
                If Not isSelected Then frmSelect.Show vbModal
            End If
        End With
    Case "II.5.C":
        With ActiveCell
            If .Row > 5 And .Row <= 386 Then
                Select Case .Column
                Case 1:    ' Sector
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SEL_YEARS")
                        .DataSetName = MSG("MSG_SELECT_YEARS")
                        .DataSource = "LST_YEARS"
                        .ReadOnly = True
                    End With
                Case 2:    ' Unit in charge
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_ADD_LOCATION_VILLAGE")
                        .DataSetName = MSG("MSG_SELECT_VILLAGE")
                        .DataSource = "tblVillage"
                        .ReadOnly = True
                    End With
                Case 3:    ' Unit in charge
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_CLIMATE_TYPE")
                        .SpecialNote = MSG("MSG_ADD_CLIMATE_TYPE")
                        .DataSource = "LST_CLIMATE_TYPE"
                    End With
                Case 4:    ' Unit in charge
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_SECTOR")
                        .DataSource = "Linhvuc"
                    End With
                Case Else
                    isSelected = True
                End Select
                If Not isSelected Then frmSelect.Show vbModal
            End If
        End With
    Case "II.5.D":
        With ActiveCell
            If .Row > 5 And .Row <= 386 Then
                Select Case .Column
                Case 1:    ' Sector
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_ADD_LOCATION_VILLAGE")
                        .DataSetName = MSG("MSG_SELECT_VILLAGE")
                        .DataSource = "tblVillage"
                        .ReadOnly = True
                    End With
                Case 2:    ' Unit in charge
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_PRODUCTS_LINE")
                        .SpecialNote = MSG("MSG_ADD_PRODUCTS_LINE")
                        .DataSource = "LST_PRODUCTS_LINE"
                    End With
                Case Else
                    isSelected = True
                End Select
                If Not isSelected Then frmSelect.Show vbModal
            End If
        End With
    Case "II.6.D":
        With ActiveCell
            If .Row > 10 And .Row <= 561 Then
                ' If the ID row is blank - get off..
                If ActiveSheet.Cells(.Row, 1) = "" Then Exit Sub
                Select Case .Column
                Case 11:    ' Sector
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_PROCURE_FORM")
                        .DataSource = "ProcureFORM"
                    End With
                Case 12:    ' Unit in charge
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_CHECK_FORM")
                        .DataSource = "CHECKFORM"
                    End With
                Case 15:
                    With frmObjectParameter
                        .DataSetName = MSG("MSG_ADD_COMPONENT")
                        .DataSource = "tblComponent"
                    End With
                Case Else
                    isSelected = True
                End Select
                With frmObjectParameter
                    .ReadOnly = True
                    .NotAllowSelection = "["
                End With
                If Not isSelected Then frmSelect.Show vbModal
            End If
        End With
    End Select
    'reset form argument value
    Dim lRet As FormArgument
    frmObjectParameter = lRet
End Sub

Function GetFolder(strPath As String, Optional FilePicker As Boolean = False, Optional FileExtension As String = "*.*") As String
    Dim fldr As FileDialog
    Dim sItem As String
    If FilePicker Then
        Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With fldr
        .Title = MSG("MSG_SELECTDATAFOLDER")
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If FilePicker Then
            .Filters.Clear
            .Filters.Add "Mirosoft Excel File", FileExtension
        End If
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Function GetDate(txtString As String) As Date
    ' This will help converting Vietnamese date to English date
    Dim Arr As Variant
    Arr = Split(Replace(txtString, "'", ""), "/")
    GetDate = DateSerial(Arr(2), Arr(1), Arr(0))
End Function

Function FormatDate(GivenDate As Date, Optional FormatType = VnDate, Optional DontSurpress As Boolean = False) As String
    ' This will override problematic date formating in Excel
    FormatDate = IIf(DontSurpress, "", "'") & Format(GivenDate, FormatType)
End Function
