Attribute VB_Name = "AddComboboxItemsModule"
Sub AddComboBoxItems()
Dim i As Range
Set ws = Worksheets(wksSettings.Name)

    With frmNewProject.cbxInstalled
        .AddItem "Yes"
        .AddItem "No"
    End With
    
        For Each i In ws.Range("ListProductionLeadTimes")
            With frmNewProject.cbxLeadTime
                .AddItem i
            End With
        Next i
        
            For Each i In ws.Range("ListMainContractor")
                With frmNewProject.cbxMainContractor
                    .AddItem i
                End With
            Next i
End Sub

