Sub setChapStyle()
    Selection.HomeKey unit:=wdLine
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.MoveDown unit:=wdLine, Count:=1
    Selection.Style = ActiveDocument.Styles("Heading 2")
End Sub

Sub setChapter()
Selection.HomeKey unit:=wdStory

    With Selection.Find
    .Forward = True
    .Wrap = wdFindStop
    .Text = "บทที่"
    .Execute
    Do While .Found
         setChapStyle
         .Execute
    Loop
    End With

End Sub
