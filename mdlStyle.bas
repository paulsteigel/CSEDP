Option Explicit

Function GenerateWordStyle(theDocument As Object, WordObj As Object) As Boolean
    ' Sets up built-in numbered list styles and List Template
    ' including restart paragraph style
    ' Run in document template during design
    ' Macro created by Margaret Aldis, Syntagma
    '
    ' Create list starting style and format if it doesn't already exist
    
    Dim strStyleName As String, tmpName  As String
    strStyleName = "Heading 1" ' the style name in this set up
    Dim strListTemplateName As String
    strListTemplateName = "SEDP_List_Template" ' the list template name in this set up
    Dim astyle As Object
        For Each astyle In theDocument.Styles
            If astyle.NameLocal = strStyleName Then GoTo Define 'already exists
        Next astyle
    ' doesn't exist
    theDocument.Styles.Add Name:=strStyleName, Type:=wdStyleTypeParagraph
Define:
    With theDocument.Styles(strStyleName)
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = wdStyleListNumber 'for international version compatibility
    End With
    With theDocument.Styles(strStyleName).ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .WidowControl = False
        .KeepWithNext = True
        .KeepTogether = True
        .OutlineLevel = wdOutlineLevelBodyText
    End With
    ' Create the list template if it doesn't exist
    Dim aListTemplate As Object
        For Each aListTemplate In theDocument.ListTemplates
            If aListTemplate.Name = strListTemplateName Then GoTo Format 'already exists
        Next aListTemplate
    ' doesn't exist
        Dim newlisttemplate As Object
        Set newlisttemplate = theDocument.ListTemplates.Add(OutlineNumbered:=True, Name:="SEDP_List_Template")
Format:
' Set up starter and three list levels - edit/extend from recorded details if required
    'Level 1
    With theDocument.ListTemplates(strListTemplateName).ListLevels(1)
        .NumberFormat = "Ph" & ChrW(7847) & "n %1:"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleUppercaseRoman
        .NumberPosition = Excel.Application.CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(0.76)
        .TabPosition = Excel.Application.CentimetersToPoints(0)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = strStyleName
    End With
    With theDocument.Styles(strStyleName)
        With .ParagraphFormat
            .LeftIndent = Excel.Application.CentimetersToPoints(0.76)
            .RightIndent = Excel.Application.CentimetersToPoints(0)
            .SpaceBefore = 12
            .SpaceAfter = 3
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphJustify
            .KeepWithNext = True
            .PageBreakBefore = False
            .FirstLineIndent = Excel.Application.CentimetersToPoints(-0.76)
            .OutlineLevel = wdOutlineLevel1
        End With
        With .Font
            .Name = "Times New Roman"
            .Size = 16
            .Bold = True
        End With
    End With

    ' Level 2
    With theDocument.ListTemplates(strListTemplateName).ListLevels(2)
        .NumberFormat = "%2."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = Excel.Application.CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(0.5)
        .TabPosition = Excel.Application.CentimetersToPoints(0.5)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 2" 'theDocument.Styles(wdStyleListNumber).NameLocal
        tmpName = "Heading 2" 'theDocument.Styles(wdStyleListNumber).NameLocal
    End With
    With theDocument.Styles(tmpName)
        With .Font
            .Name = "Times New Roman"
            .Size = 14
            .Bold = True
        End With
    End With
    
    With theDocument.ListTemplates(strListTemplateName).ListLevels(3)
        .NumberFormat = "%2.%3."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = Excel.Application.CentimetersToPoints(0.5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1)
        .TabPosition = Excel.Application.CentimetersToPoints(1)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 3" 'theDocument.Styles(wdStyleListNumber2).NameLocal
        tmpName = "Heading 3" 'theDocument.Styles(wdStyleListNumber2).NameLocal
    End With
    With theDocument.Styles(tmpName)
        With .Font
            .Name = "Times New Roman"
            .Size = 13
            .Bold = True
        End With
    End With
    
    With theDocument.ListTemplates(strListTemplateName).ListLevels(4)
        .NumberFormat = "%2.%3.%4."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.5)
        .TabPosition = Excel.Application.CentimetersToPoints(1.5)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 4" 'theDocument.Styles(wdStyleListNumber3).NameLocal
        tmpName = "Heading 4" 'theDocument.Styles(wdStyleListNumber3).NameLocal
    End With
    With theDocument.Styles(tmpName)
        With .Font
            .Name = "Times New Roman"
            .Size = 12
            .Bold = True
            .Underline = True
        End With
    End With
    With theDocument.ListTemplates(strListTemplateName).ListLevels(5)
        .NumberFormat = "%5."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.5)
        .TabPosition = Excel.Application.CentimetersToPoints(1.5)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 5" 'theDocument.Styles(wdStyleListNumber4).NameLocal
        tmpName = "Heading 5" 'theDocument.Styles(wdStyleListNumber4).NameLocal
    End With
    With theDocument.Styles(tmpName)
        With .Font
            .Name = "Times New Roman"
            .Italic = True
            .Bold = True
        End With
    End With
    With theDocument.ListTemplates(strListTemplateName).ListLevels(6)
        .NumberFormat = "%6."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleNumberInCircle
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.5)
        .TabPosition = Excel.Application.CentimetersToPoints(1.5)
        .ResetOnHigher = True
        .StartAt = 1
        With .Font
            .Bold = True
        End With
        .LinkedStyle = "Heading 6" 'theDocument.Styles(wdStyleListNumber5).NameLocal
    End With
    With theDocument.ListTemplates(strListTemplateName).ListLevels(7)
        .NumberFormat = ""
        .LinkedStyle = ""
    End With
    With theDocument.ListTemplates(strListTemplateName).ListLevels(8)
        .NumberFormat = ""
        .LinkedStyle = ""
    End With
    With theDocument.ListTemplates(strListTemplateName).ListLevels(9)
        .NumberFormat = ""
        .LinkedStyle = ""
    End With
    
    '===Bullet & Normal
    With theDocument.Styles("Normal")
        With .ParagraphFormat
            .LeftIndent = Excel.Application.CentimetersToPoints(0)
            .RightIndent = Excel.Application.CentimetersToPoints(0)
            .SpaceBefore = 3
            .SpaceAfter = 3
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = WordObj.Application.LinesToPoints(1.1)
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = Excel.Application.CentimetersToPoints(1.27)
            .OutlineLevel = wdOutlineLevelBodyText
        End With
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .NoSpaceBetweenParagraphsOfSameStyle = False
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "Normal"
    End With
           
    With theDocument
        If Not StyleExist(theDocument, "TieudeVanban") Then .Styles.Add "TieudeVanban"
        With .Styles("TieudeVanban")
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .Font.Bold = True
            With .ParagraphFormat
                .FirstLineIndent = 0
                .TabStops.Add Position:=Excel.Application.CentimetersToPoints(2.86), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
                .TabStops.Add Position:=Excel.Application.CentimetersToPoints(10.16), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
            End With
        End With
        If Not StyleExist(theDocument, "TieudeKehoach") Then .Styles.Add "TieudeKehoach"
        With .Styles("TieudeKehoach")
            .Font.Name = "Times New Roman"
            .Font.Size = 16
            .Font.Bold = True
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .SpaceBefore = 18
                .SpaceAfter = 12
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = WordObj.Application.LinesToPoints(1.1)
                .FirstLineIndent = 0
            End With
        End With
        If Not StyleExist(theDocument, "Diemnhan") Then .Styles.Add "Diemnhan"
        With .Styles("Diemnhan")
            With .ParagraphFormat
                .LeftIndent = WordObj.Application.CentimetersToPoints(1.6)
                .FirstLineIndent = WordObj.Application.CentimetersToPoints(-0.6)
            End With
            .NextParagraphStyle = wdStyleNormal 'for international version compatibility
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .Font.Bold = True
        End With
        If Not StyleExist(theDocument, "Bullet_type1") Then .Styles.Add "Bullet_type1"
        With .Styles("Bullet_type1")
            With .ParagraphFormat
                .LeftIndent = WordObj.Application.CentimetersToPoints(1.6)
                .FirstLineIndent = WordObj.Application.CentimetersToPoints(-0.6)
            End With
            .NextParagraphStyle = wdStyleNormal 'for international version compatibility
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .Font.Bold = False
        End With
        If Not StyleExist(theDocument, "HamucKehoach") Then .Styles.Add "HamucKehoach"
        With .Styles("HamucKehoach")
            .Font.Name = "Times New Roman"
            .Font.Size = 15
            .Font.Bold = True
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = Excel.Application.CentimetersToPoints(6.03)
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .FirstLineIndent = 0
            End With
        End With
        BulletText theDocument, "Diemnhan"
        ' add page number here
        .Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
    End With
    Exit Function
errHandler:
    WordObj.Quit
    GenerateWordStyle = True
End Function

Private Sub BulletText(sDoc As Object, LinkObj As String)
    Dim myList As Object

    ' Add a new ListTemplate object
    Set myList = sDoc.ListTemplates.Add

    With myList.ListLevels(1)
        .NumberFormat = ChrW(254)
        .TrailingCharacter = wdTrailingTab
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.6)
        .TabPosition = Excel.Application.CentimetersToPoints(1.6)
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = LinkObj
        ' The following sets the font attributes of
        ' the "bullet" text
        With .Font
            .Bold = False
            .Name = "Wingdings"
            .Size = 12
        End With
    End With
End Sub

Private Function StyleExist(DocObj As Object, StlName As String) As Boolean
    Dim MyStl As Object, StlObjName As String
    On Error GoTo errHandler
    Set MyStl = DocObj.Styles(StlName)
    StlObjName = MyStl.NameLocal
    StyleExist = True
errHandler:
End Function
