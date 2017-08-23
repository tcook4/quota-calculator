VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton2_Click()
'Commission Calculator

'Compensation amounts
Dim level1 As Long: level1 = 0
Dim level2 As Long: level2 = 0
Dim level3 As Long: level3 = 0

'Compensation rates
Dim rate1 As Double: rate1 = Cells(3, "G").Value
Dim rate2 As Double: rate2 = Cells(4, "G").Value
Dim rate3 As Double: rate3 = Cells(5, "G").Value

'Multipliers
Dim new_product As Double: new_product = Worksheets("Rates and Multipliers").Cells(2, 2).Value
Dim new_support As Double: new_support = Worksheets("Rates and Multipliers").Cells(3, 2).Value
Dim new_subscription As Double: new_subscription = Worksheets("Rates and Multipliers").Cells(4, 2).Value
Dim new_training As Double: new_training = Worksheets("Rates and Multipliers").Cells(5, 2).Value
Dim existing_product As Double: existing_product = Worksheets("Rates and Multipliers").Cells(2, 3).Value
Dim existing_support As Double: existing_support = Worksheets("Rates and Multipliers").Cells(3, 3).Value
Dim existing_subscription As Double: existing_subscription = Worksheets("Rates and Multipliers").Cells(4, 3).Value
Dim existing_training As Double: existing_training = Worksheets("Rates and Multipliers").Cells(5, 3).Value

'Totals and temp vaues
Dim sales_total As Long
Dim last_total As Long
Dim last_value As Long
Dim temp As Long
Dim sales_total_1 As Long
Dim sales_total_2 As Long
Dim sales_paid As Long

'Quota amounts
Dim quota As Long
quota = Cells(4, "B").Value

Dim quota_from_new As Long
Dim quota_from_existing As Long
Dim quota_from_subscription As Long

'Parsing values
Dim ws As Worksheet
Dim rw As Range
Dim RowCount As Integer

'Conditionals
Dim comp_level_1 As Boolean: comp_level_1 = True
Dim comp_level_2 As Boolean: comp_level_2 = False
Dim comp_level_3 As Boolean: comp_level_3 = False

Set sh = Worksheets("Entry Form")
    
sales_total = 0
RowCount = 1

'Start with 0$ in our commission fields
Cells(17, "B").Value = level1
Cells(18, "B").Value = level2
Cells(19, "B").Value = level3

'Traverse sale list and add to bucket
For Each rw In sh.Rows

    'Exit if we're on an empty cell
    If sh.Cells(rw.row, 1).Value = "" Then
        Exit For
    End If
    
    'New customers
    If sh.Cells(RowCount, "A") = "New" Then
        If sh.Cells(RowCount, "B") = "Product" Then
            sales_total = sales_total + ((sh.Cells(RowCount, "C").Value + sh.Cells(RowCount, "E").Value) * new_product)
            quota_from_new = quota_from_new + ((sh.Cells(RowCount, "C").Value + sh.Cells(RowCount, "E").Value) * new_product)
        ElseIf sh.Cells(RowCount, "B") = "DNS Edge" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * new_subscription)
            quota_from_new = quota_from_new + (sh.Cells(RowCount, "C").Value * new_subscription)
        ElseIf sh.Cells(RowCount, "B") = "Threat Protection" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * new_subscription)
            quota_from_new = quota_from_new + (sh.Cells(RowCount, "C").Value * new_subscription)
        ElseIf sh.Cells(RowCount, "B") = "BlueCat Private Cloud" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * new_subscription)
            quota_from_new = quota_from_new + (sh.Cells(RowCount, "C").Value * new_subscription)
        ElseIf sh.Cells(RowCount, "B") = "Training" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * new_training)
            quota_from_new = quota_from_new + (sh.Cells(RowCount, "C").Value * new_training)
        ElseIf sh.Cells(RowCount, "B") = "Enterprise Support" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * new_support)
            quota_from_new = quota_from_new + (sh.Cells(RowCount, "C").Value * new_support)
    End If
    
    'Existing Customers
    ElseIf sh.Cells(RowCount, "A") = "Existing" Then
        If sh.Cells(RowCount, "B") = "Product" Then
            sales_total = sales_total + ((sh.Cells(RowCount, "C").Value + sh.Cells(RowCount, "E").Value) * existing_product)
            quota_from_existing = quota_from_existing + ((sh.Cells(RowCount, "C").Value + sh.Cells(RowCount, "E").Value) * existing_product)
        ElseIf sh.Cells(RowCount, "B") = "DNS Edge" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * existing_subscription)
            quota_from_subscription = quota_from_subscription + (sh.Cells(RowCount, "C").Value * existing_subscription)
        ElseIf sh.Cells(RowCount, "B") = "Threat Protection" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * existing_subscription)
            quota_from_subscription = quota_from_subscription + (sh.Cells(RowCount, "C").Value * existing_subscription)
        ElseIf sh.Cells(RowCount, "B") = "BlueCat Private Cloud" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * existing_subscription)
            quota_from_subscription = quota_from_subscription + (sh.Cells(RowCount, "C").Value * existing_subscription)
        ElseIf sh.Cells(RowCount, "B") = "Training" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * existing_training)
            quota_from_subscription = quota_from_subscription + (sh.Cells(RowCount, "C").Value * existing_training)
        ElseIf sh.Cells(RowCount, "B") = "Enterprise Support" Then
            sales_total = sales_total + (sh.Cells(RowCount, "C").Value * existing_support)
            quota_from_existing = quota_from_existing + (sh.Cells(RowCount, "C").Value * existing_support)
        End If
    End If
    
    'Save our last sale value
    last_value = sales_total - last_total
    
    'level1 calc
    If comp_level_1 = True Then
        If sales_total >= quota Then
            comp_level_2 = True
            comp_level_1 = False
            level1 = rate1 * quota
            Cells(17, "B").Value = level1
            sales_total = sales_total - quota
            
            'check if we need to bump straight to 3
            If (quota_from_subscription + quota_from_new) >= quota * 0.5 Then
                comp_level_3 = True
                comp_level_2 = False
            End If
        Else
            level1 = rate1 * sales_total
            Cells(17, "B").Value = level1
        End If
    End If
    
    'level2 calc
    If comp_level_2 = True And comp_level_3 = False Then
        If (quota_from_subscription + quota_from_new) >= quota * 0.5 Then
            comp_level_3 = True
            comp_level_2 = False
            
            temp = last_value - ((quota_from_subscription + quota_from_new) - (quota * 0.5))
            sales_total = last_value - temp
            
            level2 = level2 + (temp * rate2)
            Cells(18, "B").Value = level2

        Else
            level2 = sales_total * rate2
            Cells(18, "B").Value = level2
        End If
    End If
    
    'level3 calc
    If comp_level_3 = True Then
        level3 = sales_total * rate3
        Cells(19, "B").Value = level3
    End If
    

    'update our totals and row
    last_total = sales_total
    
    RowCount = RowCount + 1
Next rw



End Sub
