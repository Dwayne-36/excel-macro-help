Attribute VB_Name = "OpenPurchaseOrdersModule"
Sub OpenPurchase(Target As Range)
Dim path As String
Dim last As Long

Set ws = ActiveSheet
On Error Resume Next
Set rng = Range("K3:K" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row)
Set Target = Target(1)
    If Not Intersect(Target, rng) Is Nothing Then
        path = "X:\Purchase Orders\Files\Purchase order.xlsm"
        Workbooks.Open Filename:=path
        Windows("Purchase Order.xlsm").Activate
        Sheets("Purchase Orders").Activate
        last = Worksheets("Purchase Orders").Cells(Worksheets("Purchase Orders").Rows.Count, "C").End(xlUp).Row + 1 ''Finding the last row of the purchase order number'''
    With Worksheets("Purchase Orders")
        .Range("A" & last).Value = Target(1).Value
        .Range("B" & last).Value = Target.Offset(0, -5).Value
    End With
        Application.EnableEvents = True
    End If
End Sub
