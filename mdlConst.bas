Option Explicit

' for Unicode caption painting
#If VBA7 Then
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcW" (ByVal hwnd As LongPtr, ByVal wMsg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Public Const WM_SETTEXT As Long = &HC

Global Const wdLineSpaceSingle = 0
Global Const wdAlignParagraphJustify = 3
Global Const wdAlignParagraphCenter = 1
Global Const wdAlignPageNumberCenter = 1
Global Const wdOutlineLevel1 = 1
Global Const wdTrailingSpace = 1
Global Const wdListNumberStyleUppercaseRoman = 1
Global Const wdUndefined = &H98967F
Global Const wdFindContinue = 1
Global Const wdReplaceAll = 2

Global Const wdListNumberStyleArabic = 0
Global Const wdListNumberStyleLowercaseLetter = 4
Global Const wdListNumberStyleNumberInCircle = &H12
Global Const wdListLevelAlignLeft = 0
Global Const wdTrailingTab = 0
Global Const wdOutlineNumberGallery = 3
Global Const wdLineSpaceMultiple = 5
Global Const wdPreferredWidthPercent = 2
Global Const wdPreferredWidthPoints = 3
Global Const wdRowHeightAtLeast = 1

Global Const wdOutlineLevelBodyText = 10
Global Const wdListNumberStyleBullet = &H17
Global Const wdStyleListNumber = &HFFFFFFCE
Global Const wdStyleListNumber2 = &HFFFFFFC5
Global Const wdStyleListNumber3 = &HFFFFFFC4
Global Const wdStyleListNumber4 = &HFFFFFFC3
Global Const wdStyleListNumber5 = &HFFFFFFC2
Global Const wdStyleNormal = &HFFFFFFFF
Global Const wdBulletGallery = 1
Global Const wdAlignTabCenter = 1
Global Const wdTabLeaderSpaces = 0
Global Const wdStyleTypeParagraph = 1

Public Enum KeyinMode   ' ChØ cho phÐp cËp nhËt ký tù ®ång kiÓu
    NumberType = 1      ' ChØ cho nhËp sè
    DateType = 2        ' NhËp kiÓu ngµy
    FormularType = 3    ' ChØ nhËp ký tù c«ng thøc
    NumberOnlyType = 4
    FreeType = 5
End Enum
Public Type LocaleSetting
    DecimalSeparator As String * 1
    GroupNumber As String * 1
    DateLocale As String * 10
End Type
Public Type FormArgument
    DataSource As String    ' Name of source range to be saved or loaded data from
    DataSetName As String   ' Name of object to be processed
    ReadOnly As Boolean     ' Define whether to lock the list
    SpecialNote As String   ' Special instruction needed
    WrapOutput As Boolean   ' Wrap output in bracket for attention
    NotAllowSelection As String ' Do not allow selection with those contained this string
    ModifyColumn As Boolean ' Tell the app to modify column data afterword
End Type
' Messages variable
Global SheetObjName As String
Global App_Title
Global ExternalLoad As Boolean
Global CurrentWorkBook As Workbook

Global AppLocale As LocaleSetting
Global ShapedLoaded As Boolean
Global frmObjectParameter As FormArgument
' for handling user event if there are any...
Global IndirectSetup As Boolean
Global AppStatus As Boolean
' for storing some temporary stuff
Global TempString As String

Sub ValidateKeycode(kCde As MSForms.ReturnInteger, Optional iNum As KeyinMode = 1)
    Select Case iNum
    Case 1:
        If kCde <> vbKeyBack Then
            If InStr("0123456789", ChrW(kCde)) = 0 Then
                If InStr(".", ChrW(kCde)) <> 0 Then
                    kCde = AscW(AppLocale.DecimalSeparator)
                Else
                    kCde = 0
                End If
            End If
        End If
    Case 2:
        If InStr("0123456789/", ChrW(kCde)) = 0 And kCde <> vbKeyBack Then kCde = 0
    Case 3:
        If InStr("()+-*/0123456789", ChrW(kCde)) = 0 And kCde <> vbKeyBack Then kCde = 0
    Case 4:
        If InStr("0123456789", ChrW(kCde)) = 0 And kCde <> vbKeyBack Then kCde = 0
    Case 5:
        ' enable free type
    End Select
End Sub

Sub NoCutAction(kCde As MSForms.ReturnInteger, ByVal ShiftKey As Integer)
    ' preventing user from pasting
    If (ShiftKey And 2) And (kCde = Asc("V")) Then kCde = 0
End Sub

Sub SetUnicodeCaption(ByVal frm As UserForm, ByVal UnicodeString As String)
    #If VBA7 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd&
    #End If
    hwnd = FindWindow("ThunderDFrame", frm.Caption)
    DefWindowProc hwnd, WM_SETTEXT, 0, StrPtr(UnicodeString)
End Sub

Function FormIsLoaded(UFName As String) As Boolean
  Dim UF As Integer
  For UF = 0 To VBA.UserForms.Count - 1
    FormIsLoaded = UserForms(UF).Name = UFName
    If FormIsLoaded Then Exit Function
  Next UF
End Function

Function HOA(chuoi As String) As String
  chuoi = Application.WorksheetFunction.Trim(chuoi)
  HOA = UCase(chuoi)
End Function

Sub RegisterAction()
    IndirectSetup = True
End Sub

Sub DeRegisterAction()
    IndirectSetup = False
End Sub
