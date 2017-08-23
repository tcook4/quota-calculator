VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
'Show our sale entry form

'Center it as best we can
With EntryForm
.Top = Int(((Application.Height / 2) + Application.Top) - (.Height / 2))
.Left = Int(((Application.Width / 2) + Application.Left) - (.Width / 2))
.Show
End With

End Sub

Private Sub CommandButton2_Click()
'Erase all and re-input the header information

'Verify with user
If MsgBox("This will erase everything! Are you sure?", vbYesNo) = vbNo Then Exit Sub

'Clear everything then re-imput header row
Sheet1.Rows(1 & ":" & Sheet1.Rows.Count).ClearContents

Sheet1.Cells(1, 1) = "Customer Type"
Sheet1.Cells(1, 2) = "Item Sold"
Sheet1.Cells(1, 3) = "Amount"
Sheet1.Cells(1, 4) = "First Year MX %"
Sheet1.Cells(1, 5) = "First Year MX $"

End Sub

Private Sub CommandButton3_Click()
'Clear the last entry in a row, unless it's the header

Dim lastrow As Long

With Sheets("Entry Form")
    lastrow = .Range("A" & .Rows.Count).End(xlUp).row
End With

If lastrow = 1 Then
    Exit Sub
End If

Sheets("Entry Form").Rows(lastrow).Clear


End Sub
