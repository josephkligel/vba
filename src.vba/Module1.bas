Sub ChangeCase()
 If TypeName(Selection) = "Range" Then
 UserForm1.Show
 Else
 MsgBox "Select a range.", vbCritical
 End If
End Sub