Sub DeleteStyle(name As String)
    Dim oStyle As Style

    For Each oStyle In ActiveDocument.Styles
        If oStyle.NameLocal = name Then
           oStyle.Delete
           Exit For
        End If
    Next oStyle
End Sub
