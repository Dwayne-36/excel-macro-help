VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call HideColumns
    Call UnhideAllRows
    Call adjust2ndRowHeight
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim rng As Range
    Dim rngUser As Range
    Dim user As String
    Dim cellColour As String
    Dim cellText As String
    Dim last As Long

    last = Sh.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Get user details from "Settings" sheet
    Set rngUser = Worksheets("Settings").Evaluate("LookupTableUsers")
    user = Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 2, False)

    ' Factory Ticks
    Set rng = Sh.Range("G3:G" & last)
    cellColour = Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 9, False)
    cellText = user
    Call FactoryTicks(Target, rng, cellColour, cellText)
    
    ' Materials Ordered
    Set rng = Sh.Range("H3:H" & last)
    cellColour = "5296274"
    Call MaterialsOrdered(Target, rng, cellColour, user)
    
    ' Release to Factory
    Set rng = Sh.Range("I3:I" & last)
    cellColour = "26350"
    Call ReleseToFactory(Target, rng, cellColour, "CNC")
    
    ' External Store
    Set rng = Sh.Range("J3:J" & last)
    cellColour = "5296274"
    Call ExternalStore(Target, rng, cellColour, user)
    
    ' Factory Ticks (N:P)
    Set rng = Sh.Range("N3:P" & last)
    cellColour = xlNone
    Call FactoryTicks(Target, rng, cellColour, ChrW(&H2713))
    
    ' Truck Booked
    Set rng = Sh.Range("S3:S" & last)
    cellColour = "26350"
    Call TruckBooked(Target, rng, cellColour)
    
    Set rng = Sh.Range("AA3:AA" & last)
    cellColour = "16751052"
    Call TruckBooked(Target, rng, cellColour)

    ' TimeCard access for specific users
    If Application.UserName = "Recption" Or Application.UserName = "Richard" Or Application.UserName = "Dwayne Keast" Then
        Call TimeCard(Target)
    End If
    
    Call OpenPurchase(Target)
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim rng As Range
    Dim last As Long

    ' Exit if the changed sheet is "Settings"
    If Sh.Name = "Settings" Then Exit Sub

    last = Sh.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Handle date-based changes
    Set rng = Sh.Range("B3:B" & last)
    If Not Intersect(Target, rng) Is Nothing And IsDate(Target.Value) Then
        Call ChangeDates(Target, DateToText(Target.Value, "mmmm yy"), last)
        Call ColorDays
    End If

    ' Assign unique ID if missing
    Set rng = Sh.Range("K3:K" & last)
    Call RandomNumber(Target, rng)
End Sub

'Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'    Static prevRow As Long
'    Dim firstCol As String: firstCol = "A"
'    Dim lastCol As String: lastCol = "AD"
'    Dim currentRow As Long
'
'    ' Check if the current Excel user is "Dwayne Keast"
'    If Application.UserName <> "Recption" Then Exit Sub
'
'    ' Get the row of the selected cell
'    currentRow = Target.Row
'
'    ' Prevent redundant updates if selection hasn't changed
'    If currentRow = prevRow Then Exit Sub
'
'    ' Clear previous highlighting
'    If prevRow > 0 Then
'        Sh.Range(firstCol & prevRow & ":" & lastCol & prevRow).Interior.ColorIndex = xlNone
'    End If
'
'    ' Highlight the new row
'    If Not Intersect(Target, Sh.Range(firstCol & ":" & lastCol)) Is Nothing Then
'        Sh.Range(firstCol & currentRow & ":" & lastCol & currentRow).Interior.Color = RGB(144, 238, 144) ' Light green
'    End If
'
'    ' Store the newly selected row
'    prevRow = currentRow
'End Sub

