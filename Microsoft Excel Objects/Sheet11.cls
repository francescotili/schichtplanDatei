VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --------------------------------------------
' MASTERWOCHEN FUNCTIONS
' Sheet for modifying masterweeks
' --------------------------------------------
Public sheet As Worksheet

Public Sub Worksheet_Activate()
  Set sheet = Sheet11
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  Dim ComboObj As OLEObject
  Dim ComboListArr As Variant
  
  If Application.CutCopyMode = 0 Then
    ' Initialize Combo Object for Autofill
    Set ComboObj = sheet.OLEObjects("TempCombo")
    With ComboObj
      .Visible = False
      .ListFillRange = "=Personal!" & Sheet9.ListObjects(GWorkerTableName).ListColumns(4).DataBodyRange.Address
      .LinkedCell = ""
    End With
        
    ' Move and visualize ComboBox if needed: only in worker range and if cell is empty
    If Not Application.Intersect(Target, Range("G7:J110, N7:Q110")) Is Nothing Then
      If Target.Areas.Count = 1 _
      And Target.Columns.Count = 1 _
      And Target.Rows.Count = 3 _
      And Target.Cells(1).Value = "" Then
        sheet.Unprotect Password:=GAdminPassword
        
        With ComboObj
          .Visible = True
          .Left = Target.Left + 1
          .Top = Target.Top + 1
          .Width = Target.Width - 1
          .Height = Target.Height - 1
          .LinkedCell = Target.Address
        End With
  
        ComboObj.Activate
        Me.TempCombo.DropDown
        
        sheet.Protect Password:=GAdminPassword
      End If
    End If
  End If
End Sub

Private Sub TempCombo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Select Case KeyCode
    Case 9 ' Tab
      changesDetectedCell.Value = True
      Application.ActiveCell.Offset(0, 1).Activate
    Case 13 ' Enter
      changesDetectedCell.Value = True
      Application.ActiveCell.Offset(1, 0).Activate
    Case 27 ' ESC
      With sheet.OLEObjects("TempCombo")
        .Object.Value = ""
        .Visible = False
      End With
      Application.ActiveCell.Activate
  End Select
End Sub
