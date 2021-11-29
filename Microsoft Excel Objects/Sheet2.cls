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
  Dim i, j As Long
  Dim weekNumber As Integer
  Dim stanzStartRow, qualiStartRow, lagerStartRow, trommStartRow, bueroStartRow, galvaStartRow, instStartRow As Integer
  Dim shiftsStartCol As Integer
  
  ' Get the week number
  weekNumber = CInt(Range("Sh2_WeekToLoadCell").Value)
  
  ' Set variable for the view
  stanzStartRow = 7
  qualiStartRow = 41
  lagerStartRow = 48
  trommStartRow = 51
  bueroStartRow = 55
  galvaStartRow = 71
  instaStartRow = 90
  shiftStartCol = 7 'G
  
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
  
  ' Stanzerei
  For i = 1 To 11
    For j = 1 To 4
      sheet.Cells((stanzStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(shiftList.weekData.Item(j).workPlace(i - 1))
    Next j
  Next i
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
