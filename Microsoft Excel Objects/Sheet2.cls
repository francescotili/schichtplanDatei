VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --------------------------------------------
' WEEK BEARBEITEN PUBLIC FUNCTIONS
' Manage methods for getting all the data for
' Sheet2 (Bearbeiten)
' --------------------------------------------
Public sheet As Worksheet

Private Sub Worksheet_Activate()
  Set sheet = ActiveWorkbook.Worksheets(GWeekModifySheetName)

End Sub

Public Sub WeekDataLoad()
  Dim shiftList As New SchichtList
  Dim i As Long
  Dim weekNumber As Integer
  
  ' Get the week number
  weekNumber = CInt(Range("Sh2_WeekToLoadCell").Value)
  
  ' Set variable for the view
'  stanzereiStartCell = "7"
'  qualitatStartCell = "38"
'  lagerStartCell = "45"
'  trommlereiStartCell = "48"
'  galvanikStartCell = "52"
'  instandsetzStartCell = "71"
'  XShiftCol = "D"
'  FShiftCol = "E"
'  SShiftCol = "F"
'  NShiftCol = "G"
  
  ' Get the week data
  shiftList.WeekRead (weekNumber)
  
  ' Set the status of week
  Select Case shiftList.weekData.Item(1).shiftStatus
    Case WeekStatus.Emtpy
      SetEmptyMode
    Case WeekStatus.Normal
      SetNormalMode
    Case WeekStatus.Custom
      SetCustomMode
  End Select
End Sub

Public Sub ToggleMode()
  Select Case Range("Sh2_WeekStatusCell").Value
    Case WeekStatus.Emtpy
      SetNormalMode
    Case WeekStatus.Normal
      SetCustomMode
    Case WeekStatus.Custom
      SetNormalMode
  End Select
End Sub

Private Sub SetNormalMode()
  Range("Sh2_WeekStatusCell").Value = WeekStatus.Normal
  Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoFalse
  Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoFalse
End Sub

Private Sub SetCustomMode()
  Range("Sh2_WeekStatusCell").Value = WeekStatus.Custom
  Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoTrue
  Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoTrue
End Sub

Private Sub SetEmptyMode()
  Range("Sh2_WeekStatusCell").Value = WeekStatus.Emtpy
  Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoFalse
  Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoFalse
End Sub
