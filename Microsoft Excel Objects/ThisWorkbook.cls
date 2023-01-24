' --------------------------------------------
' WORKBOOK FUNCTIONS
' Functions and subs that work throughtout the
' file, like at opening, savig, etc etc
' --------------------------------------------
Private Sub Workbook_Open()
  ' Bug in Excel
  ' The Worksheet_Activate event does not fire for the sheet that is active
  ' when the workbook is opened, so we call it explicitely at the Workbook
  ' opening. Make sure that Worksheet_Activate() is declared as Public.
  ' This function will ignore Worksheets without Public Activate() subs.
  On Error Resume Next
  Call Sheet2.Worksheet_Activate ' Week bearbeiten
  Call Sheet5.Worksheet_Activate ' Absence bearbeiten
  Call Sheet6.Worksheet_Activate ' ÜB Tag
  Call Sheet7.Worksheet_Activate ' ÜB Woche
  On Error GoTo 0
End Sub

