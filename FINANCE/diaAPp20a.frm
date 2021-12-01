VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp20a
   BorderStyle = 1 'Fixed Single
   Caption = "Vendor 1099 Audit"
   ClientHeight = 2925
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6090
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2925
   ScaleWidth = 6090
   Begin VB.TextBox txtTyp
      Height = 315
      Left = 1920
      TabIndex = 2
      Tag = "3"
      Top = 2430
      Width = 375
   End
   Begin VB.TextBox txtAmt
      Height = 315
      Left = 1920
      TabIndex = 1
      Tag = "1"
      Top = 1800
      Width = 975
   End
   Begin VB.ComboBox cmbVnd
      Height = 315
      Left = 1920
      Sorted = -1 'True
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Vendors"
      Top = 960
      Width = 1555
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 4920
      TabIndex = 6
      Top = 360
      Width = 1095
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaAPp20a.frx":0000
         Style = 1 'Graphical
         TabIndex = 4
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaAPp20a.frx":018A
         Style = 1 'Graphical
         TabIndex = 3
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 4920
      TabIndex = 5
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 7
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaAPp20a.frx":0308
      PictureDn = "diaAPp20a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 3960
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 2925
      FormDesignWidth = 6090
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 8
      ToolTipText = "Show System Printers"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 450
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaAPp20a.frx":0594
      PictureDn = "diaAPp20a.frx":06DA
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 4
      Left = 2520
      TabIndex = 15
      Top = 2445
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Vendor Type"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 14
      Top = 2475
      Width = 1305
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 2
      Left = 3600
      TabIndex = 13
      Top = 960
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Amount Requiring 1099"
      Height = 405
      Index = 1
      Left = 120
      TabIndex = 12
      Top = 1800
      Width = 1305
   End
   Begin VB.Label lblNme
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1920
      TabIndex = 11
      Top = 1320
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Vendor"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 10
      Top = 960
      Width = 1065
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 9
      Top = 0
      Width = 2760
   End
End
Attribute VB_Name = "diaAPp20a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaAPp20a -
'
' Notes:
'
' Created: 1/17/04 (JCW)
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub


Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
      Fillcombo
   End If
   MouseCursor 0
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaAPp20a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Fillcombo()
   FillVendors Me
   cmbVnd = ""
   cmbVnd_LostFocus
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me
End Sub

Private Sub cmbVnd_LostFocus()
   If Trim(cmbVnd) <> "" And Trim(UCase(cmbVnd)) <> "ALL" Then
      cmbVnd = CheckLen(cmbVnd, 10)
      FindVendor Me
   Else
      cmbVnd = "ALL"
      lblNme = "***Multiple Vendors Selected.***"
   End If
End Sub

Private Sub PrintReport()
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   SetMdiReportsize MdiSect
   MdiSect.crw.ReportFileName = sReportPath & "finap20a.rpt"
   
   sSql = "not Isnull({ChksTable.CHKPOSTDATE}) "
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestedBy='Requested By: " & Secure.UserInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='Vendor 1099 Audit'"
   
   If Trim(txtTyp) <> "" Then
      MdiSect.crw.Formulas(3) = "Title2='For Vendor Type: " & txtTyp & "'"
      sSql = sSql & " and {VndrTable.VETYPE} = '" & txtTyp & "' "
   Else
      MdiSect.crw.Formulas(3) = "Title2='For All Vendor Types'"
   End If
   
   If Trim(cmbVnd) <> "ALL" Then
      MdiSect.crw.Formulas(4) = "Vendor='Vendor: " & cmbVnd & "'"
      sSql = sSql & " and {VndrTable.VEREF} = '" & Compress(cmbVnd) & "' "
   Else
      MdiSect.crw.Formulas(4) = "Vendor='Vendor: All'"
   End If
   
   MdiSect.crw.Formulas(5) = "BegDate = cdate('" & Format(Now, "1/1/yy") & "')"
   MdiSect.crw.Formulas(6) = "EndDate = cdate('" & Format(Now, "12/31/yy") & "')"
   
   MdiSect.crw.Formulas(7) = "AmountRequired = " & Val(txtAmt)
   MdiSect.crw.Formulas(8) = "Title3='Amount Requiring 1099: " & Val(txtAmt) & "'"
   
   MdiSect.crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 2)
End Sub
