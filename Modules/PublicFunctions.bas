Attribute VB_Name = "PublicFunctions"
' --------------------------------------------
' MODULE FOR GLOBAL FUNCTIONS
' Here are functions used throughout the Excel
' file. They are available everywhere
' --------------------------------------------

Public Sub CleanHistory()
  Set historyList = New history
  historyList.Clean
  saveHistory str_historyCleaned
End Sub

Public Sub MitarbeiterManage()
  Dim window As New Mitarbeiter_Manage
  With window
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
  End With
End Sub

Public Sub ResetDatabase()
  Dim form As New ResetFile

  ' Open form for reset
  With form
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
  End With
End Sub

Public Function saveHistory(Optional eventName As String = str_unspecifiedHistory)
  ' Define a new History Entry
  Set newHistoryEntry = New history
  
  ' Update history
  newHistoryEntry.eventName = eventName
  newHistoryEntry.Save
End Function

Public Function columnToLetter(lngCol As Long) As String
  Dim vArr
  vArr = Split(Cells(1, lngCol).Address(True, False), "$")
  columnToLetter = vArr(0)
End Function

Public Sub test()
  MsgBox "Click!"
End Sub

Public Sub notYetReady()
  MsgBox str_notYetReady
End Sub

Public Sub notActive()
  MsgBox str_notActive
End Sub

Public Function IsInArray(stringToFind As String, dataArray As Variant) As Boolean
  For Each dataItem In dataArray
    If CStr(dataItem) = stringToFind Then
      IsInArray = True
      Exit Function
    End If
  Next
  IsInArray = False
End Function

Public Function StringFormat(ByVal mask As String, ParamArray tokens()) As String
  Dim i As Long
  
  For i = LBound(tokens) To UBound(tokens)
    mask = Replace(mask, "{" & i & "}", tokens(i))
  Next
  
  StringFormat = mask
End Function

Public Function importMasterWeek()
  Sheet2.MasterWeekImport
End Function

