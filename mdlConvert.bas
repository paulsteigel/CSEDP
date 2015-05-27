Option Explicit

'=============================================
'Global variable
Global Const AppTitleTxt = "Code Table Toolkit"

'Input parameter
Global CodeSource As String            ' Keep Source Codetable name
Global CodeDestination As String       ' Keep Destination Code table name
Global IsUpperText As Boolean          ' Whether convert to Upper case
Global IsLowerText As Boolean          ' Whether convert to Lower case
Global AutoCodeDetect As Boolean       ' Whether detecting codetable automatically - not support yet
Global FontOpt() As CodeTable          ' Keep proccessed font operator to save time loading

Global ArrayLoaded As Boolean          ' For avoiding several loaded of a large memory text

'=============================================
Private FromCharType As Boolean        ' Get the type of input char
Private ToCharType As Boolean          ' Get the type of output char
Private FrObj As Variant               ' Keep the input vowel list
Private ToObj As Variant               ' Keep the output vowel list
Private FromFontCode As Variant
Private ToFontCode As Variant
Private OldFontName As String          ' Keep previously processed Cell font

Private ActOverriden As Boolean        ' Tell whether later conversion should overriding predefined value
'=============================================
' Add in actions
'---------------------------------------------

Sub ConvertRange(CnvRange As Range)
    ' This procedure will help converting current selected range to Unicode
    Dim txtRet As String        'Keeping converted value
    Dim SetObjFont As Boolean     'Marking whether error occured or not
    
    Dim TimeObj
    TimeObj = Timer
    ' If no activesheet - get off
    If Not TestSuccess(CnvRange) Then Exit Sub
    
    'Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    Dim MyCell As Range
    
    Dim CurrentCodeSource As CodeTable, CodeSourceEx As String
    Dim oldStatusBar, AllCellCount As Long, i As Long
    Dim CurrentCode As CodeTable, FontAddUpBegin As String, FontAddUpEnd As String, CurFntString As String
    Dim t As String, MyNewFont As String
    Dim MsgProcess As String, MsgFinished As String, MsgSecond As String, MsgCell As String

    MsgProcess = MSG("MSG_PROCESS")
    MsgFinished = MSG("MSG_FINISHED_CONVERT")
    MsgSecond = MSG("MSG_SECOND")
    MsgCell = MSG("MSG_CELL")
    
    ' Initialize the input value
    InitDestination
    
    ' Load the source object if neccessary
    If Not AutoCodeDetect Then
        ' set initial value for conversion
        InitSource
        
        ' set the current code value right now
        CurrentCode = GetCodetable(CodeSource)
    End If
    CodeSourceEx = CodeSource
    
    ' Search Cell by cell and convert to anything we selected previously
    oldStatusBar = Application.StatusBar
    Application.DisplayStatusBar = True
    AllCellCount = CnvRange.Cells.Count
    
    ' Hacked 221209 - code refactor
    ' Now we have the destination code and the input value
    
    For i = 1 To AllCellCount
        Set MyCell = CnvRange.Cells(i)
        ' A test for cell error - if error - then skip
        If Not ErrorCell(MyCell) Then
            ' Since Code auto detection may overide all later settings
            ' Routine is following:
            ' 1. Detect source code
            
            If AutoCodeDetect Then
                ' try to get the previous source code to save time - simply compare font of the cell
                ' This part would help to decide what font is that
                If IsFontValid(OldFontName, MyNewFont, MyCell) Then
                    If MyNewFont <> "" Then
                        ' Now we have to set the code source due to a new font name is found
                        CurrentCode = GetSourceCodeTable(MyNewFont)
                        CodeSource = CurrentCode.t_1CodeName
                    End If
                    ' check if the destination and source are the same
                    If CodeDestination <> CodeSource Then
                        ' try to populate the fontlist with [] stuff
                        'OldFontName = LCase(CurrentCode.t_5FontConversion) - old method
                        OldFontName = CurrentCode.t_5FontConversion
                        ' now reload the source text
                        InitSource
                        GoTo SameCodeNo
                    Else
                        ' Reset the old font name
                        OldFontName = ""
                        GoTo SameCodeYes
                    End If
                Else
                    ' This cell is not good, mark it and resume next row
                    MarkCell MyCell
                    GoTo StepEnd
                End If
            End If
            ' Autodetect code will just help finding source/ destination should alway be
            ' the same - in case no sign code - then only one destination is set
SameCodeNo:
            ' So far, we should do the conversion here
            ' If we set to remove tone sign - things become so easy except vni and viqr
            ' set destination code to this
            If CurrentCode.t_1CodeName <> "" Then
                ' Now convert
                ' In fact/ if no sign code has been assigned previously, this part will do the conversion
                ' so - next time - be aware of this and try to avoid conversion again
                txtRet = ConvertText(MyCell.Value, FrObj, ToObj, FromCharType, ToCharType)
                ' a hack to prevent automatic formular stuff: 10 09 09
                If Left(txtRet, 1) = "=" Then txtRet = "'" & txtRet
            Else
                ' can not find the source code - mark the cell
                MarkCell MyCell, True
                ' just go to the next loop
                GoTo StepEnd
            End If
            ' just tell the font engine to start later
            SetObjFont = True
SameCodeYes:
            ' Previous conversion may not take place so we have to be sure that value is properly passed on this process
            If txtRet = "" Then txtRet = MyCell.Value
           
            ' now change case
            Select Case CodeDestination
            Case "Unicode", "VNI", "VIQR": ' furious hack 10 09 09
                If IsUpperText Then
                    txtRet = UCase(txtRet)
                ElseIf IsLowerText Then
                    txtRet = LCase(txtRet)
                ElseIf LCase(MyCell.Font.Name) Like LCase(CurrentCode.t_6FontUpperCase) Then
                    txtRet = UCase(txtRet)
                End If
            Case Else
                ' Now we have to find - are there any upper case text in the proccessed string
                ' this should be updated soon as we may be able to guess any upper text so that destination font can be applied
                ' hence this may reiterate through all text length, it may remarkable reduce speed
                If IsUpperText Then
                    GetFontAddUp GetCodetable(CodeDestination), FontAddUpBegin, FontAddUpEnd
                    SetObjFont = True
                ElseIf IsLowerText Then
                    ' can not do anything with this - sorry - please call Mr. Unikey
                    ' there is a way for that but I am not good at doing it by the moment!
                Else
                    If CheckUpperCase(txtRet) Then GetFontAddUp CurrentCode, FontAddUpBegin, FontAddUpEnd
                End If
            End Select
            ' Return applicable Font used
            If SetObjFont Then
                SetCellFont MyCell, FontAddUpBegin, FontAddUpEnd
            End If
            ' now parsing value to cell
            SetCellValue MyCell, txtRet
        End If
StepEnd:
        txtRet = ""
        SetObjFont = False
        ' Display the progress bar in the status line
        Application.StatusBar = MsgProcess & " " & SheetObjName & " " & i & "/" & AllCellCount & " " & MsgCell & "!"
    Next
    Application.StatusBar = MsgFinished & " " & CStr(Timer - TimeObj) & " " & MsgSecond & "!"
    
    Beep
    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.StatusBar = oldStatusBar
    ' Better cleanup all stuff so that we will be fine
    FrObj = ""
    ToObj = ""
    FromFontCode = ""
    ToFontCode = ""
    
    ' reapply current codesource and reset code counter
    CodeSource = CodeSourceEx
    OldFontName = ""
End Sub

Private Sub SetCellValue(CellObj As Range, CellValue As String)
    On Error Resume Next
    CellObj.Value = CellValue
End Sub

Private Sub SetCellFont(CellObj As Range, fPrefix As String, fSuffix As String)
    On Error Resume Next
    CellObj.Font.Name = fPrefix & FindAlterFont(CellObj.Font.Name) & fSuffix
End Sub

Private Sub MarkCell(theCell As Range, Optional ErrorCell As Boolean = False)
    ' For marking cell with problem
    On Error Resume Next
    If ErrorCell Then
        theCell.Comment.Delete
        theCell.AddComment "Can not find source code"
    Else
        theCell.Comment.Delete
        theCell.AddComment "Cell has problem"
    End If
    Error.Clear
    Resume Next
End Sub

Private Function IsFontValid(OldFont As String, NewFont As String, CurrentCellObj As Range) As Boolean
    ' This function will help clearing out all problem with font issue
    Dim tFont As String
    
    On Error GoTo errHandler
    'tFont = "[" & LCase(CurrentCellObj.Font.Name) & "]" - old method
    tFont = "[" & CurrentCellObj.Font.Name & "]"
    NewFont = ""
    
    IsFontValid = True
    
    If InStr(OldFont, tFont) <= 0 Then
        NewFont = CurrentCellObj.Font.Name
    End If
errHandler:
End Function

Private Sub GetFontAddUp(MyCode As CodeTable, FontPrefix As String, FontSuffix As String)
    If InStr(MyCode.t_6FontUpperCase, "NONE") <= 0 Then
        If Left(MyCode.t_6FontUpperCase, 1) = "*" Then
            FontSuffix = Replace(MyCode.t_6FontUpperCase, "*", "")
            FontPrefix = ""
        End If
        If Right(MyCode.t_6FontUpperCase, 1) = "*" Then
            FontSuffix = ""
            FontPrefix = Replace(MyCode.t_6FontUpperCase, "*", "")
        End If
    End If
End Sub

Private Function CheckUpperCase(InputString As String) As Boolean
    ' This function will try to see whether input string is not by mistake set in upper case
    ' There would be alot of problem so I am dumb by saying it is good or not.
    ' like I may do TEn -> no Idea whether this must use capital font or not
    ' but TEn Toi -> should likely be capital font - so I would try to see whether - occurence is 2 then
    ' use capital font
    Dim stCounter As Long, i As Long
    For i = Len(InputString) To 1 Step -1
        If Asc(Mid(InputString, i, 1)) <= 90 And Asc(Mid(InputString, i, 1)) >= 65 Then
            stCounter = stCounter + 1
        End If
        If stCounter = 2 Then
            CheckUpperCase = True
            Exit For
        End If
    Next
End Function

Private Function TestSuccess(InputValue) As Boolean
    Dim MyRange As Long
    On Error GoTo errHanlder
    MyRange = InputValue.Cells.Count
    TestSuccess = True
errHanlder:
End Function

Private Sub InitDestination()
    ' Get the source vowels list into DocVowels
    ' This acctually - will be loaded just once everytime code conversion is activated -
    ' It is a bit awkward as excel may consume energy but - it's good to clear all
    ' variable!
    Dim ToVowels As CodeTable
    
    ' Get the destination vowels list into DocVowels
    ToVowels = GetVowelList(CodeDestination, ToCharType)
    
    ' get the array of destination vowels
    ToObj = Split(ToVowels.t_2VowelList, "/")
    
    ' get the final font name set
    ToFontCode = Split(ToVowels.t_5FontConversion, "/")
End Sub

Private Sub InitSource(Optional isAutoCodeDetection As Boolean = False)
    ' Get the source vowels list into DocVowels
    ' This acctually - will be loaded just once everytime code conversion is activated -
    ' It is a bit awkward as excel may consume energy but - it's good to clear all
    ' variable!
    Dim FromVowels As CodeTable
   
    FromVowels = GetVowelList(CodeSource, FromCharType)
    
    ' get the final font name set
    FromFontCode = Split(FromVowels.t_5FontConversion, "/")
    
    ' get the array of source vowels
    FrObj = Split(FromVowels.t_2VowelList, "/")
End Sub

Function GetVowelList(ByVal mCond As String, MultiChar As Boolean) As CodeTable
    ' This part is for returning the vowels of a specific code page
    MultiChar = False
    ' Detect whether the code is multichar or single char
    GetVowelList = GetCodetable(mCond)
    If Len(GetVowelList.t_2VowelList) > 267 Then MultiChar = True
End Function

Private Function ConvertText( _
    TextToConvert As String, _
    FrObj As Variant, _
    ToObj As Variant, _
    mFrType As Boolean, _
    mToType As Boolean) As String
    
    'Routine for getting the vowel list of the selected text
    Dim i, j, k, ProcessedList() As String, ReserveList() As String
    Dim RptText As String
    ReDim ProcessedList(133)
    ReDim ReserveList(133)
    If mFrType Then
        For i = 0 To UBound(FrObj)
            If Len(FrObj(i)) = 1 Then
                ' process it later or it may cause wrong conversion
                ReserveList(k) = i
                k = k + 1
            Else
                If InStr(TextToConvert, FrObj(i)) <> 0 Then
                    ' Replace the occurence of search string with number
                    ProcessedList(j) = i
                    TextToConvert = Replace(TextToConvert, FrObj(i), "[[" & ProcessedList(j) & "]]")
                    j = j + 1
                End If
            End If
        Next
        If k > 0 Then
            ReDim Preserve ReserveList(k - 1)
            For i = 0 To UBound(ReserveList)
                If InStr(TextToConvert, FrObj(ReserveList(i))) <> 0 Then
                    ' Replace the occurence of search string with number
                    ProcessedList(j) = ReserveList(i)
                    TextToConvert = Replace(TextToConvert, FrObj(ReserveList(i)), "[[" & ProcessedList(j) & "]]")
                    j = j + 1
                End If
            Next
        End If
    Else
        For i = 0 To UBound(FrObj)
            If InStr(TextToConvert, FrObj(i)) <> 0 Then
                ' Replace the occurence of search string with number
                ProcessedList(j) = i
                TextToConvert = Replace(TextToConvert, FrObj(i), "[[" & ProcessedList(j) & "]]")
                j = j + 1
            End If
        Next
    End If
    On Error GoTo errHandle
    ReDim Preserve ProcessedList(j - 1)
    ' now just simple replace all stuff
    For i = 0 To UBound(ProcessedList)
        TextToConvert = Replace(TextToConvert, "[[" & ProcessedList(i) & "]]", ToObj(ProcessedList(i)))
    Next
    ConvertText = TextToConvert
    Exit Function
errHandle:
    If IsEmpty(j) Or IsNull(j) Then
        ' nothing to do with this text
        ConvertText = TextToConvert
    End If
End Function

Private Function FindAlterFont(iFontName As String) As String
    Dim i As Long
    If iFontName = "" Then
        iFontName = FromFontCode(0)
        GoTo iCont
    End If
    On Error GoTo iCont
    For i = LBound(FromFontCode) To UBound(FromFontCode)
        'If LCase(iFontName) Like LCase(FromFontCode(i)) & "*" Then
        If iFontName Like FromFontCode(i) & "*" Then
            FindAlterFont = ToFontCode(i)
            Exit For
        End If
    Next
iCont:
    If FindAlterFont <> "" Then Exit Function
    FindAlterFont = ToFontCode(0)
End Function

Private Function GetAlterFont(iTxtFontString As Variant) As Variant
    ' We have the range of single text, now break them to array
    Dim iTbl As Variant
    iTbl = Split(iTxtFontString, "/")
    GetAlterFont = iTbl
End Function

' Conversion stuff
Function GetSourceCodeTable(Optional intFontName As String = "") As CodeTable
    ' This function will try to get the equipvalent codetable from the built font code
    ' anyway - there no efficient way to get this dont precisely - I try to make
    ' the best matched font to convert
    ' Check the object Font first and try to loop through the font array!
    ' We kneed an array to keep font recognizer
    ' This will only activate in acase that autocodedetection is checked
    ' There are bug found by bop - nguyen in font detection procedure!
    ' In fact I may like to develop a beter font detection system that help verifying character to character...!
    ' But it seems a bit hard indeed so still use this without doing converting font name to upper case.
    Dim i As Long
    Dim FoundCode As CodeTable
    If intFontName = "" Then Exit Function
    ' see if FontOpt is set or not
    If Not ArrayLoaded Then
        ReDim FontOpt(UBound(CodeArray))
    End If
    
    ' now build the entire font list
    If Not ArrayLoaded Then
        For i = 0 To UBound(CodeArray)
            FontOpt(i) = GetCodetable(CodeArray(i))
            FontOpt(i).t_5FontConversion = PopulateFontList(FontOpt(i).t_5FontConversion)
            FontOpt(i).t_1CodeName = CodeArray(i)
        Next
    End If
    For i = 0 To UBound(CodeArray)
        'If IsContainedText(LCase(intFontName), LCase(FontOpt(i).t_4FontRecognizer)) Then old method
        If IsContainedText(intFontName, FontOpt(i).t_4FontRecognizer) Then
            ' this seems a good font
            FoundCode = FontOpt(i)
            Exit For
        End If
    Next
    ' Mark the loaded array
    ArrayLoaded = True
    
    GetSourceCodeTable = FoundCode
    ' Now we will try to guest the code by input string
    If FontOpt(i).t_1CodeName = "" Then
        ' found no code matched - so we have to mark this object
    End If
End Function

Private Function PopulateFontList(IntTxt As String) As String
    Dim OutString As String ', i As Long
'    Dim tblTxt As Variant
'    tblTxt = Split(IntTxt, "/")
    'For i = 0 To UBound(tblTxt)
'        outString = outString & "/" & "[" & tblTxt(i) & "]"
    'Next
    
    OutString = "[" & Replace(IntTxt, "/", "]/[") & "]"
    PopulateFontList = OutString
End Function

Private Function IsContainedText(TxtIn As String, txtCompare As String) As Boolean
    Dim tblArr As Variant, i As Long
    tblArr = Split(txtCompare, "/")
    For i = 0 To UBound(tblArr)
        'If LCase(txtIn) Like LCase(tblArr(i)) & "*" Then ' old method
        If TxtIn Like tblArr(i) & "*" Then
            IsContainedText = True
            Exit For
        End If
    Next
End Function

Private Function ErrorCell(inCell As Range) As Boolean
    On Error GoTo ErrorHandler
    ErrorCell = True
    If Not IsEmpty(inCell) And Not inCell.HasFormula And Trim(inCell) <> "" Then ErrorCell = False
ErrorHandler:
End Function
