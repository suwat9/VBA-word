Sub Arabic2Thai(mode As Boolean)
Dim n1, n2   As Integer
If mode Then
   'Mode = true , Arabic to Thai
   n1 = 48
   n2 = 240
Else
   'Mode = false, Thai to Arabic
    n1 = 240
    n2 = 48
End If

For i = 0 To 9
With Selection.Find
.Text = Chr(n1 + i)
.Replacement.Text = Chr(n2 + i)
.Wrap = wdFindContinue
End With
Selection.Find.Execute Replace:=wdReplaceAll
Next
End Sub
