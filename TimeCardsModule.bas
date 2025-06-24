Attribute VB_Name = "TimeCardsModule"
Function TimeCard(ByVal Target As Range)
On Error Resume Next

Set rng = Range("F3:F" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row)

    If Target.Count > 1 Then Exit Function
        If Not Intersect(Target, rng) Is Nothing Then
                Application.EnableEvents = False
                Cancel = True 'stop the edit mode
            With Target
                frmProjectTimes.Show
            End With
            Application.EnableEvents = True
    End If
End Function
Sub AddProject()
Dim ws As Worksheet
Windows("Time Card.xlsx").Activate
Set ws = Worksheets("ProjectTimes")

last = lastRow("ProjectTimes") + 1

    With frmProjectTimes
        Jobnumber = .lblJobNumber
        project = .tbProjectName
        hours = .tbHours
        whiteBoards = .tbWhiteBoards
        colourBoards = .tbColourBoards
        cabinetQty = .lblCabinetQty
    End With
        With ws
            .Range(ReturnLastRowAddress).Offset(0, -1) = Jobnumber
            .Range(ReturnLastRowAddress).Offset(0, 13) = cabinetQty
            .Range(ReturnLastRowAddress) = project
                If InStr(1, frmProjectTimes.cbxDepartment.Text, "Cut") > 0 Then
                   If frmProjectTimes.cbxDepartment = "Cut colour" Then
                       .Range(ReturnLastRowAddress).Offset(-1, dpt) = colourBoards
                       .Range(ReturnLastRowAddress).Offset(-1, dpt - 1) = hours
                   Else
                       .Range(ReturnLastRowAddress).Offset(-1, dpt) = whiteBoards
                       .Range(ReturnLastRowAddress).Offset(-1, dpt - 1) = hours
                   End If
                Else
                    .Range(ReturnLastRowAddress).Offset(-1, dpt - 1) = hours
                End If
         End With
End Sub
Sub UpdateProject()
Dim ws As Worksheet
Windows("Time Card.xlsx").Activate
Set ws = Worksheets("ProjectTimes")

        With frmProjectTimes
                hours = .tbHours
                whiteBoards = .tbWhiteBoards
                colourBoards = .tbColourBoards
        End With

        With ws
             If .Range(ReturnProjectAddress) > 0 Then
                .Range(ReturnProjectAddress) = .Range(ReturnProjectAddress) + hours
             Else
                .Range(ReturnProjectAddress) = hours
             End If
                          
                If InStr(1, frmProjectTimes.cbxDepartment.Text, "Cut") > 0 Then
                    If frmProjectTimes.cbxDepartment = "Cut colour" Then
                        If .Range(ReturnProjectAddress).Offset(0, 1) > 0 Then
                            .Range(ReturnProjectAddress).Offset(0, 1) = .Range(ReturnProjectAddress).Offset(0, 1) + colourBoards
                        Else
                            .Range(ReturnProjectAddress).Offset(0, 1) = colourBoards
                        End If
                    Else
                        If .Range(ReturnProjectAddress).Offset(0, 1) > 0 Then
                            .Range(ReturnProjectAddress).Offset(0, 1) = .Range(ReturnProjectAddress).Offset(0, 1) + whiteBoards
                        Else
                            .Range(ReturnProjectAddress).Offset(0, 1) = whiteBoards
                        End If
                    End If
                End If
         End With
End Sub
Sub AddProjectName()
Dim prodAddress As String
frmProjectTimes.tbProjectName = Range(ProAddress("Dispatch V2.xlsm", 5)).Value
frmProjectTimes.lblCabinetQty = Range(ProAddress("Dispatch V2.xlsm", 7)).Value
End Sub
Function ProAddress(WB As String, os As String)
Dim WorkSheetName As String
 Windows(WB).Activate
Set ws = ActiveSheet
WorkSheetName = ActiveSheet.Name
    Dim rngX As Range
        Set rngX = ws.Range("F1:F" & lastRow(WorkSheetName)).Find(frmProjectTimes.lblJobNumber, lookat:=xlPart)
            If Not rngX Is Nothing Then
                ProAddress = rngX.Offset(0, os).Address
            End If
End Function
Function FindProject()
    Dim rngX As Range
    Windows("Time Card.xlsx").Activate
    Set ws = Worksheets("ProjectTimes")

last = lastRow("ProjectTimes")

        Set rngX = ws.Range("A1:A" & last).Find(frmProjectTimes.lblJobNumber, lookat:=xlPart)
            If Not rngX Is Nothing Then
                FindProject = True
            Else
                FindProject = False
            End If
            
End Function
Function ReturnProjectAddress()
    Dim rngX As Range
        Set rngX = Worksheets("ProjectTimes").Range("A1:A" & lastRow("ProjectTimes")).Find(frmProjectTimes.lblJobNumber, lookat:=xlPart)
            If Not rngX Is Nothing Then
                ReturnProjectAddress = rngX.Offset(0, dpt).Address
            End If
End Function

Function ReturnLastRowAddress()
    With Sheets("ProjectTimes")
        ReturnLastRowAddress = .Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Address
    End With
End Function
Function dpt()
department = frmProjectTimes.cbxDepartment.Value
    If department = "Cut colour" Then
        dpt = 3
    ElseIf department = "Cut white" Then
        dpt = 5
    ElseIf department = "Edge" Then
        dpt = 7
    ElseIf department = "Pre" Then
        dpt = 8
    ElseIf department = "Ass" Then
        dpt = 9
    Else     '''Dezignatek'''
        dpt = 10
    End If
End Function
Sub LoadComboDepartment()
    With frmProjectTimes.cbxDepartment
        .AddItem "Cut colour"
        .AddItem "Cut white"
        .AddItem "Dezignatek"
        .AddItem "Edge"
        .AddItem "Pre"
        .AddItem "Ass"
    End With
End Sub
Sub LoadListProjects()
    Dim i As Long
    Dim j As Long
    Dim Temp As Variant

    Windows("Dispatch V2.xlsm").Activate
Set ws = ActiveSheet
    last = lastRow(ActiveSheet.Name)
    frmProjectTimes.lbProjects.Clear

    For Each p In ws.Range("F3:F" & last)
        If p.Value <> vbNullString Then frmProjectTimes.lbProjects.AddItem p
    Next p
    
    With frmProjectTimes.lbProjects
        For i = 0 To .ListCount - 2
            For j = i + 1 To .ListCount - 1
                If .List(i) > .List(j) Then
                    Temp = .List(j)
                    .List(j) = .List(i)
                    .List(i) = Temp
                End If
            Next j
        Next i
    End With
End Sub
Function TimeCardValidate()
    If frmProjectTimes.cbxDepartment = vbNullString Then
        MsgBox "Please Select a Department.", vbExclamation, "Department!"
        frmProjectTimes.cbxDepartment.SetFocus
        TimeCardValidate = "No"
            Exit Function
    End If

        If frmProjectTimes.cbxDepartment = "Cut colour" Then
            If frmProjectTimes.tbHours = vbNullString Then
                MsgBox "Please enter hours.", vbExclamation, "Hours!"
                frmProjectTimes.cbxDepartment.SetFocus
                TimeCardValidate = "No"
                    Exit Function
            ElseIf frmProjectTimes.tbColourBoards = vbNullString Then
                MsgBox "Please enter board quantity.", vbExclamation, "Colour Board Qty!"
                frmProjectTimes.cbxDepartment.SetFocus
                TimeCardValidate = "No"
                    Exit Function
            End If
        End If
        
        If frmProjectTimes.cbxDepartment = "Cut white" Then
            If frmProjectTimes.tbHours = vbNullString Then
                MsgBox "Please enter hours.", vbExclamation, "Hours!"
                frmProjectTimes.cbxDepartment.SetFocus
                TimeCardValidate = "No"
                    Exit Function
            ElseIf frmProjectTimes.tbWhiteBoards = vbNullString Then
                MsgBox "Please enter board quantity.", vbExclamation, "White Board!"
                frmProjectTimes.cbxDepartment.SetFocus
                TimeCardValidate = "No"
                    Exit Function
            End If
        End If
        
        If frmProjectTimes.tbHours = vbNullString Then
                MsgBox "Please enter hours.", vbExclamation, "Hours!"
                frmProjectTimes.tbHours.SetFocus
                TimeCardValidate = "No"
            Exit Function
        End If
     
    TimeCardValidate = "Yes"
    
End Function
Sub CheckIfWorkBookOpen(myWB As String)
Dim WB As Workbook

    For Each WB In Workbooks
        If WB.Name = myWB Then
            WB.Close saveChanges:=False
                Exit Sub
        End If
    Next WB
    
End Sub
Sub enable()
Application.EnableEvents = True
End Sub
