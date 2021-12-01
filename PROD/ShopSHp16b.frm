VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp16a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Sales Backlog with MO Status"
   ClientHeight = 1860
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 5430
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 1860
   ScaleWidth = 5430
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdHlp
      Appearance = 0 'Flat
      Height = 250
      Left = 0
      Picture = "ShopSHp16b.frx":0000
      Style = 1 'Graphical
      TabIndex = 11
      TabStop = 0 'False
      ToolTipText = "Subject Help"
      Top = 0
      UseMaskColor = -1 'True
      Width = 250
   End
   Begin VB.CheckBox optDsc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2040
      TabIndex = 2
      Top = 1440
      Value = 1 'Checked
      Width = 735
   End
   Begin VB.ComboBox txtBeg
      Height = 315
      Left = 2040
      TabIndex = 1
      Tag = "4"
      Top = 360
      Width = 1095
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 3960
      TabIndex = 7
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 3960
      TabIndex = 4
      Top = 360
      Width = 1095
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "ShopSHp16b.frx":07AE
         Style = 1 'Graphical
         TabIndex = 5
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "ShopSHp16b.frx":092C
         Style = 1 'Graphical
         TabIndex = 6
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin VB.ComboBox cmbDivision
      Height = 315
      Left = 2040
      Sorted = -1 'True
      TabIndex = 0
      Tag = "8"
      ToolTipText = "Contains Customers With Allocations"
      Top = 780
      Width = 1095
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4140
      Top = 1020
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 1860
      FormDesignWidth = 5430
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 285
      Index = 5
      Left = 240
      TabIndex = 10
      Top = 1200
      Width = 1695
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Extended Descriptions"
      Height = 285
      Index = 3
      Left = 240
      TabIndex = 9
      Tag = " "
      Top = 1440
      Width = 1725
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "As of Date"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 8
      Tag = " "
      Top = 360
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Division"
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 3
      Top = 780
      Width = 1425
   End
End
Attribute VB_Name = "ShopSHp16a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/28/05 Changed date handling
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Sub cmbDivision_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbDivision = CheckLen(cmbDivision, 10)
   For iList = 0 To cmbDivision.ListCount - 1
      If cmbDivision = cmbDivision.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbDivision = cmbDivision.List(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "select distinct SODIVISION from SohdTable order by SODIVISION"
   LoadComboBox cmbDivision, 1
   If cmbDivision.ListCount > 0 Then
      cmbDivision = cmbDivision.List(0)
   Else
      cmbDivision = "No SOs for any division"
      cmbDivision.ForeColor = ES_RED
   End If
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHp16a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sAsOfDate As String
   
   MouseCursor 13
   On Error Resume Next
   sCust = Compress(cmbDivision)
   If Not IsDate(txtBeg) Then
      sAsOfDate = "1995,01,01"
   Else
      sAsOfDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   
   MouseCursor 13
   On Error GoTo DiaErr1
   SetMdiReportsize MDISect
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "AsOf='" & txtBeg & "'"
   MDISect.Crw.Formulas(2) = "RequestBy = 'Requested By: " & sInitials & "'"
   sCustomReport = GetCustomReport("slebl08")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = ""
   MDISect.Crw.Formulas(3) = "ExtendedDesc='" & optDsc.Value & "'"
   MDISect.Crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub














Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtBeg = Now
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "sh16", Trim(Str(optDsc.Value))
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh06", sOptions)
   If Len(sOptions) Then
      optDsc.Value = Val(sOptions)
   Else
      optDsc.Value = vbChecked
   End If
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 3 Then
      txtBeg = CheckDate(txtBeg)
   Else
      txtBeg = "ALL"
   End If
   
End Sub
