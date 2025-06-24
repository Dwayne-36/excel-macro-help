Attribute VB_Name = "EnterProjectModule"
Sub EnterProject(WorkSheetN As String, Optional DispatchOrAllJobs As String)
Dim ws As Worksheet
Dim RowNumber As Integer
Dim LeadTime As Integer

Set ws = Sheets(WorkSheetN)
If Not DispatchOrAllJobs = "All Jobs" Then
    Set ProductionLead = Names("LookupTableProductionLeadTimes").RefersToRange
End If
last = lastRow(WorkSheetN)

        With frmNewProject
                col_A = .lblWeekNumber1                                                                           '''Week no'''
                col_B = .tbDispatchDate                                                                           '''Dispatch date'''
                If Not DispatchOrAllJobs = "All Jobs" Then
                    col_C = AddSubtractDays(.tbDispatchDate, Application.WorksheetFunction.VLookup(.cbxLeadTime.Value, ProductionLead, 3, False), "Subtract", "d-mmm")      '''Production date'''
                    col_D = AddSubtractDays(.tbDispatchDate, Application.WorksheetFunction.VLookup(.cbxLeadTime.Value, ProductionLead, 4, False), "Subtract", "d-mmm")      '''Detail date'''
                End If
                col_E = .tbDispatchDate                                                                           '''Design date'''
                col_F = .lblJobNumber1
                Col_J = .cbxMainContractor                                                                        '''Main contractor'''
                col_K = .tbProjectName                                                                            '''Project Name
                col_L = .tbProjectColour                                                                          '''Project Colour'''
                col_M = .tbQty                                                                                    '''Quantity'''
                col_R = .cbxInstalled                                                                             '''Installed'''
                col_S = .tbFreight                                                                                '''Freight'''
                col_T = .tbBenchtopSupplier                                                                       '''Benchtop supplier'''
                col_U = .tbBenchtopColour                                                                         '''Benchtop colour'''
                col_V = .tbInstaller                                                                              '''Installer'''
                col_W = .tbComment                                                                                '''Comment'''
                col_X = .tbDeliveryAddress                                                                        '''Delivery addres'''
                col_Y = .tbPhone                                                                                  '''Builder Phone'''
                col_Z = .tbM3                                                                                     '''Meters squared'''
                col_AA = .tbAmount                                                                                '''Amount'''
                col_AB = .tbOrderNumber                                                                           '''Order number'''
                col_AC = .tbDateOrdered                                                                           '''Date Ordered'''
                col_AD = .cbxLeadTime.Value                                                                       '''Lead time'''
        End With
        
                If Not DispatchOrAllJobs = "All Jobs" Then
                    RowNumber = last + 2
                Else
                    RowNumber = last + 1
                End If

        With ws
                .Range("A" & RowNumber) = col_A
                .Range("B" & RowNumber) = col_B
                If Not DispatchOrAllJobs = "All Jobs" Then
                    .Range("C" & RowNumber) = col_C
                    .Range("D" & RowNumber) = col_D
                End If
                .Range("E" & RowNumber) = col_E
                .Range("F" & RowNumber) = col_F
                .Range("J" & RowNumber) = Col_J
                    If Col_J = "J Scene" Or Col_J = "A1 Chch" Or Col_J = "A1 Crom" Then
                        .Range("J" & RowNumber).Offset(0, -9).Interior.Color = 255
                    End If
                .Range("K" & RowNumber) = col_K
                .Range("L" & RowNumber) = col_L
                .Range("M" & RowNumber) = col_M
                .Range("R" & RowNumber) = col_R
                .Range("S" & RowNumber) = col_S
                .Range("T" & RowNumber) = col_T
                .Range("U" & RowNumber) = col_U
                .Range("V" & RowNumber) = col_V
                .Range("W" & RowNumber) = col_W
                .Range("X" & RowNumber) = col_X
                .Range("Y" & RowNumber) = col_Y
                .Range("Z" & RowNumber) = col_Z
                .Range("AA" & RowNumber) = col_AA
                .Range("AB" & RowNumber) = col_AB
                .Range("AC" & RowNumber) = col_AC
                .Range("AD" & RowNumber) = col_AD
            End With
End Sub
