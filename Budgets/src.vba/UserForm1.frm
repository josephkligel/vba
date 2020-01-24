Private Sub OptionButton1_Click()

End Sub

Private Sub CanelButton_Click()
Unload UserForm1
End Sub

Private Sub OKButton_Click()
 Dim WorkRange As Range
 Dim cell As Range
'Process only text cells, no formulas
 On Error Resume Next
 Set WorkRange = Selection.SpecialCells _
 (xlCellTypeConstants, xlCellTypeConstants)
'Upper case
 If OptionUpper Then
 For Each cell In WorkRange
 cell.Value = UCase(cell.Value)
 Next cell
 End If
' Lower case
 If OptionLower Then
 For Each cell In WorkRange
 cell.Value = LCase(cell.Value)
 Next cell
 End If
' Proper case
 If OptionProper Then
 For Each cell In WorkRange
 cell.Value = Application. _
 WorksheetFunction.Proper(cell.Value)
  Next cell
 End If
 Unload UserForm1
End Sub