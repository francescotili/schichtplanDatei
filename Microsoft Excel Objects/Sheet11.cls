' --------------------------------------------
' MASTERWOCHEN FUNCTIONS
' Sheet for modifying masterweeks
' --------------------------------------------
Public sheet As Worksheet

Public Sub Worksheet_Activate()
  Set sheet = Sheet11
  ResetConditionalFormatting
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
      Application.ActiveCell.Offset(0, 1).Activate
      saveHistory str_sh11_history
    Case 13 ' Enter
      Application.ActiveCell.Offset(1, 0).Activate
      saveHistory str_sh11_history
    Case 27 ' ESC
      With sheet.OLEObjects("TempCombo")
        .Object.Value = ""
        .Visible = False
      End With
      Application.ActiveCell.Activate
  End Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  ' Monitor changes in Masterwoche registerkarte
  saveHistory str_sh11_history
End Sub

Private Sub ResetConditionalFormatting()
  Dim bigRange, stanzRange, qualiRange, lagerRange, bueroRange, galvaRange, instRange, elektroRange As Range

  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword

  ' Define ranges for every section
  Set bigRange = Range("G7:J105,N7:Q105")
  Set stanzRange = Range("G7:J39,N7:Q39")
  Set qualiRange = Range("G40:J45,N40:Q45")
  Set lagerRange = Range("G46:J51,N46:Q51")
  Set bueroRange = Range("G52:J66,N52:Q66")
  Set galvaRange = Range("G67:J84,N67:Q84")
  Set instRange = Range("G85:J102,N85:Q102")
  Set elektroRange = Range("G103:J105,N103:Q105")

  ' Delete preexistent conditional formatting
  bigRange.FormatConditions.Delete

  ' Add italic style for extern employees
  With bigRange
    .FormatConditions.Add Type:=xlExpression, Formula1:="=NOT(ISNUMBER(MATCH(G7;INDIRECT(""MitarbeiterList[Abkürzung]"");0)))"
    With .FormatConditions(1)
      .Font.Bold = False
      .Font.Italic = True
      .StopIfTrue = False
    End With
  End With

  ' Add conditional formatting for Stanzerei
  With stanzRange
    ' Add fill for employees
    .FormatConditions.Add Type:=xlExpression, Formula1:="=G7<>"""""
    With .FormatConditions(2)
      With .Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.499984740745262
      End With
      With .Borders
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
      End With
      With .Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
      End With
      .StopIfTrue = False
    End With
  End With

  ' Add conditional formatting for Qualität
  With qualiRange
    ' Add fill for employees
    .FormatConditions.Add Type:=xlExpression, Formula1:="=G40<>"""""
    With .FormatConditions(2)
      With .Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = -0.499984740745262
      End With
      With .Borders
        .LineStyle = xlContinuous
        .ThemeColor = 8
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
      End With
      With .Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
      End With
      .StopIfTrue = False
    End With
  End With

  ' Add conditional formatting for Lager
  With lagerRange
    ' Add fill for employees
    .FormatConditions.Add Type:=xlExpression, Formula1:="=G46<>"""""
    With .FormatConditions(2)
      With .Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.499984740745262
      End With
      With .Borders
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With .Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
      End With
      .StopIfTrue = False
    End With
  End With

  ' Add conditional formatting for Büro
  With bueroRange
    ' Add fill for employees
    .FormatConditions.Add Type:=xlExpression, Formula1:="=G52<>"""""
    With .FormatConditions(2)
      With .Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.749961851863155
      End With
      With .Borders
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.749961851863155
        .Weight = xlThin
      End With
      With .Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99481185338908E-02
      End With
      .StopIfTrue = False
    End With
  End With

  ' Add conditional formatting for Galvanik
  With galvaRange
    ' Add fill for employees
    .FormatConditions.Add Type:=xlExpression, Formula1:="=G67<>"""""
    With .FormatConditions(2)
      With .Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = -0.499984740745262
      End With
      With .Borders
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With .Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
      End With
      .StopIfTrue = False
    End With
  End With

  ' Add conditional formatting for Instandsetzung
  With instRange
    ' Add fill for employees
    .FormatConditions.Add Type:=xlExpression, Formula1:="=G85<>"""""
    With .FormatConditions(2)
      With .Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
      End With
      With .Borders
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With .Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
      End With
      .StopIfTrue = False
    End With
  End With

  ' Add conditional formatting for Elektriker
  With elektroRange
    ' Add fill for employees
    .FormatConditions.Add Type:=xlExpression, Formula1:="=G103<>"""""
    With .FormatConditions(2)
      With .Font
        .Bold = True
        .Italic = False
        .Color = -8840844
        .TintAndShade = 0
      End With
      With .Borders
        .LineStyle = xlContinuous
        .Color = -8840844
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With .Interior
        .PatternColorIndex = xlAutomatic
        .Color = 16309751
        .TintAndShade = 0
      End With
      .StopIfTrue = False
    End With
  End With

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub