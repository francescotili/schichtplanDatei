VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
