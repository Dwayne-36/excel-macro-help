Attribute VB_Name = "MainModule"
Sub ShowNewProjectsForm()
frmNewProject.Show vbModeless
End Sub
Public Sub Main(WorkSheetN As String)
last = lastRow(WorkSheetN)
LastC = lastCol(WorkSheetN, "Letter")

Call FilterBlankRows(last, WorkSheetN)


Call DeleteBlankRows(last, WorkSheetN, 2)
Call Sort(last, WorkSheetN, LastC)
Call InsertBlankRows(last, WorkSheetN)

last = lastRow(WorkSheetN)

Call Boarders(WorkSheetN, LastC, last, "Yes")
Call ColourBlankRows(WorkSheetN, last, LastC)
Call WeeklyTotal(WorkSheetN, last)

End Sub


