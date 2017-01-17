Sub setChapStyle()
    Selection.HomeKey unit:=wdLine
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.MoveDown unit:=wdLine, Count:=1
    Selection.Style = ActiveDocument.Styles("Heading 2")
End Sub

Sub setChapterReplace()
Selection.HomeKey unit:=wdStory

    With Selection.Find
    .Forward = True
    .Wrap = wdFindStop
    .Text = "บทที่"
    .Replacement.Text = "Chapter"
    .Execute Replace:=wdReplaceOne, Forward:=True
    Do While .Found
         setChapStyle
         .Execute Replace:=wdReplaceOne, Forward:=True
    Loop
    End With

End Sub
