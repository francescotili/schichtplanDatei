VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declare the variables of the History object
Public eventName As String
Private currentUser As New User
Private sheet As Worksheet
Private table As ListObject

Private Sub Init()
    eventName = ""
End Sub

Public Function Save()
  Dim entry As Variant
  Dim i As Integer

  Set sheet = ActiveWorkbook.Worksheets(GHistorySheetName)
  Set table = sheet.ListObjects.Item(GHistoryTableName)
    
  ' Create array Value for History
  entry = Array( _
    CDbl(Now()), _
    currentUser.name, _
    eventName, _
    currentUser.location, _
    currentUser.computer _
  )
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
    
  ' Add new row at first position
  table.ListRows.Add Position:=1
  ' Populate the row with data
  For i = LBound(entry) To UBound(entry)
    table.DataBodyRange(1, (i + 1)).Value = entry(i)
  Next i
  
  ' Enable screen Updating and protection
  sheet.Protect Password:=GAdminPassword
  Application.ScreenUpdating = True
End Function
