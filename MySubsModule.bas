Attribute VB_Name = "MySubsModule"
'Option Explicit
Global ws As Worksheet
Sub DeleteBlankRows(last, WorkSheetN, First)

    Set ws = Sheets(WorkSheetN)      '<--alter as needed
        With ws
            .Range("A" & First & ":A" & last + 3).SpecialCells(xlCellTypeVisible).EntireRow.Delete '<--alter row to exclude headers

            On Error GoTo 0         'turn error checking back on
            .AutoFilterMode = False

        End With
End Sub
Sub FilterBlankRows(last, WorkSheetN)

    Set ws = Sheets(WorkSheetN)      '<--alter as needed
        With ws
            .Range("A1:A" & last).AutoFilter Field:=1, Criteria1:="="     '<--alter column as needed
            
        On Error Resume Next    'in case there are none to delete
        
        End With
End Sub
Sub InsertBlankRows(last, WorkSheetN)

Dim rng As Range, cel As Range
    Set ws = Sheets(WorkSheetN)
        With ws
            Set rng = Range("B2:B" & last)
                For Each cel In rng
                    If cel.Offset(-1, 0).Value <> "" And cel.Offset(-1, 0).Value <> cel.Value Then
                        Cells(cel.Row, 1).Resize(1, 30).Insert
                    End If
                Next cel
        End With
End Sub
Sub Sort(last, WorkSheetN, LastC)
Dim SortOne As String
Dim SortTwo As String
Dim SortThree As String
Dim rngUser As Range
Set ws = Sheets(WorkSheetN)
Set rngUser = Worksheets("Settings").Evaluate("LookupTableUsers")

user = Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 3, False)

    With ws.Sort
    
        .SortFields.Clear
        
        SortOne = user
'        SortTwo = "E2:E"
        SortThree = "J2:J"
        
        .SortFields.Add Key:=Range(SortOne & last), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                
'        .SortFields.Add Key:=Range(SortTwo & last), _
'                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SortFields.Add Key:=Range(SortThree & last), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                              
        .SetRange Range("A2:" & LastC & last)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
                   
    End With
End Sub

Public Sub WeeklyTotal(WorkSheetN, last)
    Dim rng As Range
    Dim cel As Range
    Dim weekTot As Integer
    Dim grandTot As Integer
    
    Set ws = Sheets(WorkSheetN)
    With ws
       
        Set rng = .Range("M3:M" & last + 1)
        
        For Each cel In rng
            'if it's not blank it's not between days
            If cel.Value <> "" Then
                weekTot = weekTot + cel.Value
            Else
                'it is blank so check if end of a week
                If Cells(cel.Row - 1, 1).Value = Cells(cel.Row + 1, 1).Value Then
                    'same week don't need to do anything
                    'but if you want to deal with individual days
                    'this is where you would do it
                Else
                    'different week need to do something
                    With cel.Offset(0, 0)
                        .Value = weekTot
                        .Interior.ColorIndex = 44
                    End With
                    'alter totals as necessary
                    grandTot = grandTot + weekTot
                    weekTot = 0
                End If
            End If
        Next cel
       
        'now you can write the grandtotal where ever you want it
        With Cells(last + 3, 13)
            .Value = grandTot
            .Interior.ColorIndex = 45
        End With
        
    End With

End Sub
Sub Boarders(WorkSheetN, Col, last, YesorNo)

Set ws = Sheets(WorkSheetN)
        With ws
                Range("A1:" & Col & last).Select
            If YesorNo = "Yes" Then
                Selection.Borders.LineStyle = xlContinuous
            Else
                Selection.Borders.LineStyle = xlNone
            End If
    End With
End Sub
Sub ColourBlankRows(WorkSheetN, last, lastCol)
Dim crng As Range, cel As Range
Dim clr As Long
Set ws = Sheets(WorkSheetN)

    With ws
        Set crng = .Range("B2:B" & last)
                'Colouring the black rows in
            For Each cel In crng
                If cel.Offset(1, 0).Value <> "" And cel.Offset(1, 0).Value <> cel.Value Then
                
                   With Cells(cel.Row, 1).Resize(1, Cells(1, Columns.Count).End(xlToLeft).Column)
                        .Interior.ColorIndex = 15
                        .Borders(xlEdgeRight).LineStyle = xlNone
                        .Borders(xlEdgeLeft).LineStyle = xlNone
                        .Borders(xlInsideVertical).LineStyle = xlNone
                        .RowHeight = 12
                   End With
                            
                End If
            Next cel
                'Colouring the last row at the bottom
                 Range("A" & last + 1 & ":" & lastCol & last + 1).Select
                    Selection.RowHeight = 12
            With Selection.Interior
                .ColorIndex = 15
            End With
    End With
End Sub

Sub ChangeDates(ByVal Target As Range, WorkSheetN, last)
On Error Resume Next
Dim rng As Range
Dim wf As WorksheetFunction
Dim LeadTime As Integer

With ws
    Set ws = Sheets(WorkSheetN)
    Set ProductionLead = Names("LookupTableProductionLeadTimes").RefersToRange
    Set wf = Application.WorksheetFunction
    Set Target = Target(1)
    Set rng = Worksheets("Settings").Range("A:A").Find(What:=DateToText(Target.Value, "dd-mmm-yy"), LookIn:=xlValues, lookat:=xlWhole)
    
    If Format(Target.Value, "ddd") = "Sat" Or Format(Target.Value, "ddd") = "Sun" Then
        MsgBox "You have selected a date that falls on the weekend. Please try again."
        Application.Undo
        Target.Select
        Exit Sub
    Else
            If Not rng Is Nothing Then
                MsgBox "You have selected a day that falls on a weekend or holiday. Please try again."
                Application.Undo
                Target.Select
                Exit Sub
            End If
    Set rng = Range("B3:B" & last)
        If Not Intersect(Target, rng) Is Nothing And IsDate(Target.Value) Then
        
        
        LeadTime = Application.WorksheetFunction.VLookup(Target.Offset(0, 28).Value, ProductionLead, 3, False)
            Target.Offset(0, -1).Value = ISOweeknum(Target.Value)
                Target.Offset(0, 1).Value = AddSubtractDays(Target.Value, LeadTime, "Subtract", "d-mmm")
                    Target.Offset(0, 3).Value = Target.Value
                        LeadTime = Application.WorksheetFunction.VLookup(Target.Offset(0, 28).Value, ProductionLead, 4, False)
                            Target.Offset(0, 2).Value = AddSubtractDays(Target.Value, LeadTime, "Subtract", "d-mmm")
        End If
    End If
End With

End Sub
Sub ClearNewProjectsForm()
With frmNewProject
                .lblWeekNumber1 = ""
                .tbDispatchDate = ""                                                                           '''Dispatch date'''
                .tbDispatchDate = ""                                                                           '''Design date'''
                .lblJobNumber1 = ""
                .cbxMainContractor = ""                                                                        '''Main contractor'''
                .tbProjectName = ""                                                                            '''Project Name
                .tbProjectColour = ""                                                                          '''Project Colour'''
                .tbQty = ""                                                                                    '''Quantity'''
                .cbxInstalled = ""                                                                             '''Installed'''
                .tbFreight = ""                                                                                '''Freight'''
                .tbBenchtopSupplier = ""                                                                       '''Benchtop supplier'''
                .tbBenchtopColour = ""                                                                         '''Benchtop colour'''
                .tbInstaller = ""                                                                              '''Installer'''
                .tbComment = ""                                                                                '''Comment'''
                .tbDeliveryAddress = ""                                                                        '''Delivery addres'''
                .tbPhone = ""                                                                                  '''Builder Phone'''
                .tbM3 = ""                                                                                     '''Meters squared'''
                .tbAmount = ""                                                                                 '''Amount'''
                .tbOrderNumber = ""                                                                            '''Order number'''
                .cbxLeadTime = ""                                                                              '''Lead time'''
        End With
End Sub
Sub UpdateJobNumber(Maincontractor As String, Jobnumber As Integer)

Application.EnableEvents = False

    JobNumberAddress = FindInNamedRange("LookupTableMainContractor", Maincontractor)
    
        Worksheets("Settings").Range(JobNumberAddress).Offset(0, 2) = Jobnumber + 1
        
Application.EnableEvents = True
    
End Sub
Sub Refresh()
Dim ws As Worksheet
Dim WorkSheetN As String
    If ActiveSheet.Name = "Settings" Then
        Exit Sub
    End If
        WorkSheetN = ActiveSheet.Name
            Set ws = Sheets(WorkSheetN)
    
                Application.ScreenUpdating = False
                    
                ws.Activate
                
                    Call Main(WorkSheetN)
                    
                Application.ScreenUpdating = True
End Sub
Sub AllJobsV2()

                Workbooks.Open Filename:="X:\Schedules\All Jobs V2.xlsx"
                Windows("All Jobs V2.xlsx").Activate
                Sheets("Data").Select
                
                Call EnterProject("Data", "All Jobs")

                ActiveWorkbook.Save
                ActiveWindow.Close
                Windows("Dispatch V2.xlsm").Activate
                
End Sub
Sub HideColumns()
Dim ws As Worksheet
Dim rngUser As Range
Set rngUser = Worksheets("Settings").Evaluate("LookupTableUsers")

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Settings" Then
            With ws
                .Columns("B:B").Hidden = Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 4, False)
                .Columns("C:C").Hidden = Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 5, False)
                .Columns("D:D").Hidden = Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 6, False)
            End With
        End If
    Next ws
End Sub
Sub UnhideAllRows()
Dim ws As Worksheet
    
        For Each ws In Worksheets
            If ws.Name = "Settings" Or ws.Name = "Template" Or ws.Name = "Remeadials" Then
                Exit Sub
            Else
                ws.Rows.EntireRow.Hidden = False
            End If
        Next ws
    
 
End Sub
'Sub HideRows()
'Dim ws As Worksheet
'Dim lastRow As Integer
'Dim actCell As Range
'Dim rngUser As Range
'Set rngUser = Worksheets("Settings").Evaluate("LookupTableUsers")
'
'Application.ScreenUpdating = False
'    For Each ws In Worksheets
'        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
'
'        For Each actCell In ws.Range("A3:A" & lastRow)
'            With actCell
'                If ws.Name = "Settings" Or ws.Name = "Remeadials" Then
'                    Exit Sub
'                Else
'                    If Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 7, False) = "YES" Then
'                        If ws.Range("J" & .Row) = "J Scene" And ws.Range("A" & .Row).Interior.Color = 255 Then
'                            .EntireRow.Hidden = True
'                         ElseIf ws.Range("J" & .Row) = "A1 Chch" And ws.Range("A" & .Row).Interior.Color = 255 Then
'                            .EntireRow.Hidden = True
'                         ElseIf ws.Range("J" & .Row) = "A1 Crom" And ws.Range("A" & .Row).Interior.Color = 255 Then
'                            .EntireRow.Hidden = True
'                         Else
'                            .EntireRow.Hidden = False
'                        End If
'                    ElseIf Application.WorksheetFunction.VLookup(Application.UserName, rngUser, 8, False) = "YES" Then
'                        If ws.Range("J" & .Row) = "J Scene" Or ws.Range("J" & .Row) = "A1 Chch" Or ws.Range("J" & .Row) = "A1 Crom" Then
'                            .EntireRow.Hidden = False
'                        Else
'                            .EntireRow.Hidden = True
'                        End If
'                    End If
'                End If
'            End With
'        Next
'    Next
'Application.ScreenUpdating = True
'End Sub
        'Select the newly inserted row
Sub SelectRow(WorkSheetN As String, FindWhat As String)
Dim rng As Range
    Set rng = Worksheets(WorkSheetN).Range("F:F").Find(What:=FindWhat, LookIn:=xlValues, lookat:=xlWhole)
        If rng Is Nothing Then
            MsgBox "Could not find the selected row. Please try again."
        Else
            Range(Range(rng.Address).Offset(0, -5).Address).Resize(, lastCol(WorkSheetN, "Number")).Activate
            ActiveWindow.ScrollRow = Selection.Row
        End If
End Sub
        'Saves a copy to the backup folder
Sub SaveBackup()
Dim datetime As String

datetime = DateToText(Now, "dd-mmm-hh-mm AM/PM")
JobPath = "X:\Schedules\Backup\" & DateToText(Now, "yyyy") & "\" & Application.UserName & " " & "Dispatch V2 " & "(" & datetime & ")"

    Application.ScreenUpdating = False
    
        ChDir "X:\Schedules\Backup\" & DateToText(Now, "yyyy")
        ActiveWorkbook.SaveCopyAs Filename:=JobPath & ".xlsm"
    
    Application.ScreenUpdating = True
End Sub
Sub adjust2ndRowHeight()
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        With ws
        .Rows(1).RowHeight = 127
        .Rows(2).RowHeight = 10
        End With
    Next ws
End Sub
