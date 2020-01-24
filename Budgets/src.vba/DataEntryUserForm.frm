
Private Sub Label4_Click()

End Sub

Private Sub CancelButton_Click()
Unload DataEntryUserForm
End Sub

Private Sub CatComboBox_Change()

End Sub

Private Sub DatePicker_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub PmtComboBox_Change()

End Sub

Private Sub ResetButton_Click()
Unload Me
DataEntryUserForm.Show
End Sub

Private Sub OKButton_Click()
    Dim NextRow As Long
    
'   Make sure Sheet1 is active
    Sheets("Transactions").Activate
    
'   Determine the next empty row
    NextRow = Application.WorksheetFunction. _
        CountA(Range("A:A"))
        
'   Transfer the date
    Cells(NextRow, 1) = DatePicker.Value
    
'   Transfer the payment type
    Cells(NextRow, 2) = PmtComboBox.Value
    
'   Transfer the description
    Cells(NextRow, 3) = DescTextBox.Text

'   Transfer the type
    Cells(NextRow, 4) = TypeComboBox.Value
    
'   Transfer the category
    Cells(NextRow, 5) = CatComboBox.Value
    
'   Transfer the amount
    Cells(NextRow, 6) = AmtTextBox.Text
    
'   Clear the controls for the next entry
    DatePicker.SetFocus
    
'   Make sure correct information is entered
    'Dim Msg As String
    'Msg = "You selected items: "
    'Msg = Msg & vbNewLine
    'Msg = Msg & "Date: " & DatePicker.Value
    'Msg = Msg & vbNewLine
    'Msg = Msg & "Pmt Type: " & PmtComboBox.Value
    'Msg = Msg & vbNewLine
    'Msg = Msg & "Type: " & TypeComboBox.Value
    'Msg = Msg & vbNewLine
    'Msg = Msg & "Category: " & CatComboBox.Value
    'Msg = Msg & vbNewLine
    'Msg = Msg & "Amount: " & AmtTextBox.Value
    'MsgBox Msg
    Unload DataEntryUserForm
End Sub
Private Sub PaymentTypeLabel_Click()

End Sub

Private Sub TypeComboBox_Change()
CatComboBox.RowSource = TypeComboBox
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
'   Fill the Date Picker
    DatePicker.Value = Date
'   Fill the combo box
    With PmtComboBox
        .AddItem "Account/Transfer"
        .AddItem "Card"
    End With
'   Select the second list item
    PmtComboBox.ListIndex = 1

' Fill the combo box
    With TypeComboBox
    .AddItem "Income"
    .AddItem "Expense"
    End With
' Select the second list item
    TypeComboBox.ListIndex = 1
End Sub