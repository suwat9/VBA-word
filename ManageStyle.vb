Sub ManageStyle(StyleName As String)
    ActiveDocument.Styles.Add name:=StyleName, Type:=wdStyleTypeParagraph
    ActiveDocument.Styles(StyleName).AutomaticallyUpdate = False
    With ActiveDocument.Styles(StyleName).Font
        .name = "+Body"
        .Size = 14
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 16
        .NameBi = "TH Sarabun New"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles(StyleName).ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles(StyleName).NoSpaceBetweenParagraphsOfSameStyle = _
        False
    ActiveDocument.Styles(StyleName).ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles(StyleName).ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles(StyleName).LanguageID = wdThai
    ActiveDocument.Styles(StyleName).NoProofing = False
    With ActiveDocument.Styles(StyleName).Frame
        .TextWrap = True
        .WidthRule = wdFrameAuto
        .HeightRule = wdFrameAuto
        .HorizontalPosition = wdFrameLeft
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
        .VerticalPosition = CentimetersToPoints(0)
        .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
        .HorizontalDistanceFromText = CentimetersToPoints(0)
        .VerticalDistanceFromText = CentimetersToPoints(0)
        .LockAnchor = False
    End With
End Sub
