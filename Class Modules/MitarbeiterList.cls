VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MitarbeiterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------------------
' MITARBEITER LIST CLASS
' Manage methods for getting all the users or
' search for specific ones
' --------------------------------------------

Public data As Collection
'Public user As Mitarbeiter
Public worker As Mitarbeiter
Private sheet As Worksheet
Private table As ListObject

Private Sub Class_Initialize()
  Set sheet = ActiveWorkbook.Worksheets(GUserSheetName)
  Set table = sheet.ListObjects.Item(GWorkerTableName)
End Sub

Public Sub Load()
  ' Get all the userTable and build a Collection
  ' of Mitarbeiter objects
  
  Dim workersData As Variant
  Set data = New Collection
  Dim i As Long
  Dim tempWorker As Mitarbeiter
  
  If Not table.DataBodyRange Is Nothing Then
    workersData = table.DataBodyRange
  
    ' Parse data in Mitarbeiter objects
    For i = LBound(workersData) To UBound(workersData)
      Set tempWorker = New Mitarbeiter
      tempWorker.Init _
        CStr(workersData(i, 1)), _
        CStr(workersData(i, 2)), _
        CStr(workersData(i, 3)), _
        CStr(workersData(i, 4)), _
        CStr(workersData(i, 5)), _
        CStr(workersData(i, 6))
    
      data.Add tempWorker
    Next i
  End If
End Sub

Public Sub Search(workerPCode As String)
  ' Get all the userTable and search for a specific
  ' user through its Personal Code
  
  Dim workersData As Variant
  Dim i As Long
  
  workersData = table.DataBodyRange
  
  ' Search for passed userPCode
  For i = LBound(workersData) To UBound(workersData)
    If workersData(i, 1) = workerPCode Then
      Set worker = New Mitarbeiter
      worker.Init _
        CStr(workersData(i, 1)), _
        CStr(workersData(i, 2)), _
        CStr(workersData(i, 3)), _
        CStr(workersData(i, 4)), _
        CStr(workersData(i, 5)), _
        CStr(workersData(i, 6))
    End If
  Next i
End Sub

Public Sub Save(newWorker As Mitarbeiter)
  ' Save or modify the passed user
  
  Dim i, j As Long
  Dim existingWorker As Integer
  Dim entry As Variant

  ' Load the user Table
  Me.Load
  existingWorker = -1
  
  ' Check if user already exists
  For i = 1 To data.Count
    If data.Item(i).personalCode = newWorker.personalCode Then
      existingWorker = i
    End If
  Next i
  
  ' Create array Value for Worker
  entry = Array( _
    newWorker.personalCode, _
    newWorker.name, _
    newWorker.surname, _
    newWorker.visName, _
    newWorker.department, _
    newWorker.vacationTotal _
  )
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  If existingWorker = -1 Then ' New user
    ' Add new row at first position
    table.ListRows.Add Position:=1
    ' Populate the row with data
    For j = LBound(entry) To UBound(entry)
      table.DataBodyRange(1, (j + 1)).Value = entry(j)
    Next j
    ' Update history
    saveHistory ("Neu Mitarbeiter (" & newWorker.personalCode & ") wurde hinzugefügt")
    ' Add the new worker in the Vacation Plan
    Dim vacationList As New AbwesenheitsList
    vacationList.WorkerAdd newWorker
  Else ' User already exists
    For j = LBound(entry) To UBound(entry)
      table.DataBodyRange(existingWorker, (j + 1)).Value = entry(j)
    Next j
    saveHistory ("Mitarbeiter-Nr. " & newWorker.personalCode & " wurde bearbeiten")
  End If
  
  ' Sort user for surname
  Dim sortColumn As Range
  Set sortColumn = Range(GWorkerTableName & "[Nachname]")
  With table.Sort
    .SortFields.Clear
    .SortFields.Add Key:=sortColumn, SortOn:=xlSortOnValues, Order:=xlAscending
    .Header = xlYes
    .Apply
  End With
  
  ' Updates pivot tables
  sheet.PivotTables("MitarbeiterList_Stanzerei").RefreshTable
  
  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub Delete(personalCode As String)
  Dim i As Long
  Dim deletingWorker As Integer

  ' Load the user Table
  Me.Load
  deletingWorker = -1
  
  ' Search for the user
  For i = 1 To data.Count
    If data.Item(i).personalCode = personalCode Then
      deletingWorker = i
    End If
  Next i

  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  If deletingWorker <> -1 Then ' Delete the row
    table.ListRows(deletingWorker).Delete
  End If
  
  ' Updates pivot tables
  sheet.PivotTables("MitarbeiterList_Stanzerei").RefreshTable

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub Reset()
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Delete all the users
  If Not table.DataBodyRange Is Nothing Then
    table.DataBodyRange.Delete
  End If
  
  ' Updates pivot tables
  sheet.PivotTables("MitarbeiterList_Stanzerei").RefreshTable

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
  
  ' Delete all user from Vacation List
  Dim vacationList As New AbwesenheitsList
  vacationList.WorkerDelete
End Sub
