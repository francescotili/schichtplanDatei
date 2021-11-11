Attribute VB_Name = "PublicFunctions"
' --------------------------------------------
' MODULE FOR GLOBAL FUNCTIONS
' Here are functions used throughout the Excel
' file. They are available everywhere
' --------------------------------------------

Sub CleanHistory()
  Set historyList = New history
  historyList.Clean
End Sub

Sub MitarbeiterManage()
  Dim window As New Mitarbeiter_Manage
  With window
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
  End With
End Sub

Public Function saveHistory(Optional eventName As String = "Unspecified event detected")
  ' Define a new History Entry
  Set newHistoryEntry = New history
  
  ' Update history
  newHistoryEntry.eventName = eventName
  newHistoryEntry.Save
End Function

Public Sub test()
  Dim test As New AbwesenheitsList
  test.GenerateYear
End Sub

