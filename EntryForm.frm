VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm 
   Caption         =   "Entry Form"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   OleObjectBlob   =   "EntryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_button_Click()

Unload Me

End Sub


Private Sub Submit_button_Click()

'Verify all fields have a value
If Customer_entry.ListIndex = -1 Then
    MsgBox "Please enter new or existing customer"
    Exit Sub
ElseIf Item_entry.ListIndex = -1 Then
    MsgBox "Please enter an item sold"
    Exit Sub
ElseIf Amount_entry.Value = "" Then
    MsgBox "Please enter a price for item sold"
    Exit Sub
ElseIf Mx_entry.Value = "" And Mx_entry_dollar.Value = "" And Item_entry.Value = "Product" Then
    MsgBox "Please enter a percentage or dollar amount for first year maintenance"
    Exit Sub
ElseIf Mx_entry.Value <> "" And Mx_entry_dollar.Value <> "" And Item_entry.Value = "Product" Then
    MsgBox "Please enter either a percent value for maintenance or dollar amount - not both"
    Exit Sub
End If

Dim emptyRow As Long

'Make Sheet1 active
Sheet1.Activate

'Determine emptyRow
emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

'Transfer information
Cells(emptyRow, 1).Value = Customer_entry.Value
Cells(emptyRow, 2).Value = Item_entry.Value
Cells(emptyRow, 3).Value = Amount_entry.Value

'Parse First Year MX Value
If Mx_entry.Value >= 1 And Mx_entry.Value <> "" Then
    Mx_entry.Value = Mx_entry.Value / 100
End If


'Generate formula for first year MX
Dim row As String
row = emptyRow
If Mx_entry.Value <> "" And Item_entry.Value = "Product" Then
    Cells(emptyRow, 4).Value = Mx_entry.Value
    Cells(emptyRow, 5).Value = Mx_entry.Value * Amount_entry.Value
ElseIf Mx_entry_dollar.Value <> "" And Item_entry.Value = "Product" Then
    Cells(emptyRow, 5).Value = Mx_entry_dollar.Value
    Cells(emptyRow, 4).Value = Mx_entry_dollar.Value / Amount_entry.Value
End If

'Check for more than one entry
If Entries_check.Value = True Then
    Amount_entry.Value = ""
    Mx_entry.Value = ""
    Mx_entry_dollar.Value = ""
Else
    Unload Me
End If

Sheet2.Activate

End Sub

Private Sub UserForm_Initialize()

'Empty CustomerType box
Customer_entry.Clear

'Fill Customer Type box
With Customer_entry
    .AddItem "New"
    .AddItem "Existing"
End With

'Empty Item Sold box
Item_entry.Clear

'Fill Item Sold
With Item_entry
    .AddItem "Product"
    .AddItem "DNS Edge"
    .AddItem "Threat Protection"
    .AddItem "BlueCat Private Cloud"
    .AddItem "Enterprise Support"
    .AddItem "Training"
    .AddItem "Other"
End With


'Empty Amount box
Amount_entry.Value = ""


'Empty Maintaince entry
Mx_entry.Value = ""
Mx_entry_dollar.Value = ""

'Uncheck Multiple box
Entries_check.Value = False

'Set Focus on Amount
Amount_entry.SetFocus

End Sub

