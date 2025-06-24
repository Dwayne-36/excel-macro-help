Attribute VB_Name = "MyFunctionsModule"
        'Returns the last used row
Function lastRow(WorkSheetName As String)
    With Sheets(WorkSheetName)
'        LastRow = .Cells.Find("*", .Cells(.Rows.Count, .Columns.Count), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
         lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    End With
End Function
        'Returns the last used column Letters
Function lastCol(WorkSheetName As String, Optional NumberOrLetter As String)
    With Sheets(WorkSheetName)
        If NumberOrLetter = "Letter" Then
            LastColNumber = Cells(1, Columns.Count).End(xlToLeft).Column
            lastCol = Split(Cells(1, LastColNumber).Address, "$")(1)    'Returns a letter
        Else
            lastCol = Cells(1, Columns.Count).End(xlToLeft).Column      'Returns a number
        End If
        
        
    End With
End Function

        'Add or subtracts days - Startdate is the date to add or subtract from. The number of days is how many days to be added or subtracted. AddOrSubtrach, Type "Add" to add or "Subtract" to subtract.
Function AddSubtractDays(StartDate As Date, NumberOfDays As Integer, AddOrSubtract As String, FormatDate As String)
Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    Set ListHolidays = Names("ListHolidays").RefersToRange
    
        If AddOrSubtract = "Add" Then
            AddSubtractDays = Format(wf.WorkDay(StartDate, NumberOfDays, ListHolidays), FormatDate)
        Else
            AddSubtractDays = Format(wf.WorkDay(StartDate, -NumberOfDays, ListHolidays), FormatDate)
        End If
End Function
        'Returns the current date
Function GetDate(Optional DateFormat As String)
   GetDate = Format(Date, DateFormat)
End Function
        'Returns the week number
Public Function ISOweeknum(ByVal v_Date As Date) As Integer
ISOweeknum = DatePart("ww", v_Date - Weekday(v_Date, 2) + 4, 2, 2)
End Function
        'Converts a date to a text
Function DateToText(DateToConvert, DateFormat)
    DateToText = Format(DateToConvert, DateFormat)
End Function
        'Markers for "Run" , "Edge" , "Assemble"
Function FactoryTicks(ByVal Target As Range, ByVal rng As Range, ByVal cellColour As String, ByVal cellText As String)
On Error Resume Next

If Target.Count > 1 Then Exit Function
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False
        Cancel = True 'stop the edit mode
        With Target
            If .Value = "" Then
                .Value = cellText
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                    With Selection.Interior
                        .Color = cellColour
                    End With
            Else
                 Selection.ClearContents
                 With Selection.Interior
                      .Pattern = xlNone
                 End With
            End If
            
           .Offset(0, 1).Activate
        End With
        Application.EnableEvents = True
    End If
End Function
Function ReleseToFactory(ByVal Target As Range, ByVal rng As Range, ByVal cellColour As String, ByVal cellText As String)
On Error Resume Next

If Target.Count > 1 Then Exit Function
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False
        Cancel = True 'stop the edit mode
            With Target
                    If Selection.Interior.Pattern <> RGB(255, 255, 255) Then
                        If .Value = "" Then
                                .Value = cellText
                        ElseIf .Value = cellText Then
                                Selection.Interior.Color = cellColour
                                .Value = "REL"
                        ElseIf .Value = "REL" Then
                                If .Value = "REL" Then
                                Selection.ClearContents
                                End If
                                With Selection.Interior
                                    .Pattern = xlNone
                                End With
                        End If
                    Else
                        If .Value = "REL" Then
                                Selection.ClearContents
                                With Selection.Interior
                                    .Pattern = xlNone
                                End With
                         End If
                    End If
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .Offset(0, 1).Activate
            End With
        Application.EnableEvents = True
    End If
End Function
Function MaterialsOrdered(ByVal Target As Range, ByVal rng As Range, ByVal cellColour As String, ByVal cellText As String)
On Error Resume Next

If Target.Count > 1 Then Exit Function
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False
        Cancel = True 'stop the edit mode
            With Target
                    If Selection.Interior.Pattern <> RGB(255, 255, 255) Then
                        If .Value = "" Then
                                .Value = cellText
                                Selection.Interior.Color = "255"
                        ElseIf .Value = cellText And Selection.Interior.Color = "255" Then
                                Selection.Interior.Color = cellColour
                                .Value = cellText
                        ElseIf .Value = cellText Then
                                If .Value = cellText Then
                                Selection.ClearContents
                                End If
                                With Selection.Interior
                                    .Pattern = xlNone
                                End With
                        End If
                    Else
                        If .Value = cellText Then
                                Selection.ClearContents
                                With Selection.Interior
                                    .Pattern = xlNone
                                End With
                         End If
                    End If
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .Offset(0, 1).Activate
            End With
        Application.EnableEvents = True
    End If
End Function
Function ExternalStore(ByVal Target As Range, ByVal rng As Range, ByVal cellColour As String, ByVal cellText As String)
On Error Resume Next

If Target.Count > 1 Then Exit Function
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False
        Cancel = True 'stop the edit mode
            With Target
                    If Selection.Interior.Pattern <> RGB(255, 255, 255) Then
                            If Selection.Interior.Color = "26350" Then
                                Selection.Interior.Color = cellColour
                            ElseIf Selection.Interior.Color = cellColour Then
                                    With Selection.Interior
                                        .Pattern = xlNone
                                    End With
                            ElseIf Selection.Interior.Color <> "26350" And Selection.Interior.Color <> cellColour Then
                                Selection.Interior.Color = "26350"
                            End If
                    End If
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlBottom
                        .Offset(1, 0).Activate
            End With
        Application.EnableEvents = True
    End If
End Function

Function TruckBooked(ByVal Target As Range, ByVal rng As Range, ByVal cellColour As String)
On Error Resume Next

If Target.Count > 1 Then Exit Function
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False
        Cancel = True 'stop the edit mode
        With Target
            If Selection.Interior.Pattern = xlNone Then
                    With Selection.Interior
                        .Color = cellColour
                    End With
            Else
                 With Selection.Interior
                      .Pattern = xlNone
                 End With
            End If
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .Offset(0, 1).Activate
        End With
        Application.EnableEvents = True
     End If
End Function
Public Function ColorDays()
    Dim FirstAddress As String
    Dim MySearch As Variant
    Dim myColor As Variant
    Dim rng As Range
    Dim i As Long
    Dim Sh As Worksheet
    
    last = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    MySearch = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", _
    "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", _
    "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52")

    myColor = Array("4", "6", "7", "8", "22", "4", "6", "7", "8", "22", "4", "6", "7", "8", "22", "4", "6", "7", "8", "22", "4", _
    "6", "7", "8", "22", "4", "6", "7", "8", "22", "4", "6", "7", "8", "22", "4", "6", "7", "8", "22", "4", "6", "7", "8", "22", _
    "4", "6", "7", "8", "22", "4", "6")

        With ActiveSheet.Range("A3:A" & last)
            
'            .Interior.ColorIndex = xlColorIndexNone

            For i = LBound(MySearch) To UBound(MySearch)

                'If you want to find a part of the rng.value then use xlPart
                'if you use LookIn:=xlValues it will also work with a
                'formula cell that evaluates to MySearch(I)             Offset(0, 1).Value = d2
                Set rng = .Find(What:=MySearch(i), _
                                After:=.Cells(.Cells.Count), _
                                lookat:=xlWhole, _
                                SearchOrder:=xlByColumns, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)

                If Not rng Is Nothing Then
                    FirstAddress = rng.Address
                    Do
                        rng.Offset(0, 4).Interior.ColorIndex = myColor(i)
                        Set rng = .FindNext(rng)
                    Loop While Not rng Is Nothing And rng.Address <> FirstAddress
                End If
            Next i
        End With
'    Next sh
End Function

Function FindInNamedRange(NamedRange As String, FindWhat As String)
Dim c As Range

    With ThisWorkbook.Names(NamedRange).RefersToRange
        Set c = .Find(FindWhat, LookIn:=xlValues)
        If c Is Nothing Then
        Else
            FindInNamedRange = (c.Address)
        End If
    End With
End Function
Function Validate()
Dim ws As Worksheet
Dim Sheet As String
Dim rng As Range
    Sheet = DateToText(frmNewProject.tbDispatchDate, "mmmm yy")
        Set ws = Sheets(Sheet)
            If frmNewProject.cbxLeadTime.Value = vbNullString Then
                MsgBox "You need to enter a the leadtime!", vbExclamation, "No Answer Found!"
                frmNewProject.cbxLeadTime.SetFocus
                Validate = "No"
                Exit Function
            End If
            If frmNewProject.tbProjectName.Text = vbNullString Then
                MsgBox "You need to enter a project Name!", vbExclamation, "No Answer Found!"
                frmNewProject.tbProjectName.SetFocus
                Validate = "No"
                Exit Function
            End If
            If frmNewProject.tbDispatchDate.Text = vbNullString Then
                MsgBox "You need to enter a dispatch date!", vbExclamation, "No Answer Found!"
                frmNewProject.tbDispatchDate.SetFocus
                Validate = "No"
                Exit Function
            End If
            If frmNewProject.cbxMainContractor.Value = vbNullString Then
                MsgBox "You need to enter a the franchise!", vbExclamation, "No Answer Found!"
                frmNewProject.cbxMainContractor.SetFocus
                Validate = "No"
                Exit Function
            End If
            If frmNewProject.tbProjectColour.Text = vbNullString Then
                MsgBox "You need to enter the project colour(s)!", vbExclamation, "No Answer Found!"
                frmNewProject.tbProjectColour.SetFocus
                Validate = "No"
                Exit Function
            End If
            If frmNewProject.tbQty = vbNullString Then
                MsgBox "You need to enter 'CABINET QTY!'", vbExclamation, "No Answer Found!"
                frmNewProject.tbQty.SetFocus
                Validate = "No"
                Exit Function
            End If
            If frmNewProject.lblDay1.Caption = "Saturday" Or frmNewProject.lblDay1.Caption = "Sunday" Then
                MsgBox "You have choesen a date that is on a weekend. Please try again."
                frmNewProject.tbDispatchDate.SetFocus
                Validate = "No"
            Exit Function
            End If
            
            Set rng = Worksheets("Settings").Range("A:A").Find(What:=DateToText(frmNewProject.tbDispatchDate, "dd-mmm-yy"), LookIn:=xlValues, lookat:=xlWhole)
            
            If Not rng Is Nothing Then
                MsgBox "You have selected a holiday. Pleas try again"
                frmNewProject.tbDispatchDate.SetFocus
                Validate = "No"
                Exit Function
            End If
            Validate = "Yes"
End Function
        'Check if sheet exists or not
Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook

    Dim Sheet As Object
    For Each Sheet In InWorkbook.Sheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
    sheetExists = False
End Function

