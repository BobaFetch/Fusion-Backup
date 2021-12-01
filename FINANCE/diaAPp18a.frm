VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp18a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Cash Disbursements By GL Account (Report)"
   ClientHeight = 5415
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 5835
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 5415
   ScaleWidth = 5835
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox chkSummary
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 32
      Top = 4560
      Width = 855
   End
   Begin VB.CheckBox chkVnd
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 29
      Top = 4920
      Width = 855
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6000
      Top = 4680
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5415
      FormDesignWidth = 5835
   End
   Begin VB.CheckBox ChkExt
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 26
      Top = 4320
      Width = 855
   End
   Begin VB.CheckBox chkComp
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 25
      Top = 4080
      Width = 855
   End
   Begin VB.ComboBox txtEndDte
      Height = 315
      Left = 1440
      TabIndex = 12
      Top = 1800
      Width = 1215
   End
   Begin VB.ComboBox txtBegDte
      Height = 315
      Left = 1440
      TabIndex = 11
      Top = 1440
      Width = 1215
   End
   Begin VB.TextBox txtEndNum
      Height = 285
      Left = 1440
      TabIndex = 10
      Tag = "1"
      Top = 840
      Width = 1095
   End
   Begin VB.TextBox txtBegNum
      Height = 285
      Left = 1440
      TabIndex = 9
      Tag = "1"
      Top = 480
      Width = 1095
   End
   Begin VB.ComboBox cmbvnd
      Height = 315
      Left = 1440
      TabIndex = 8
      Top = 2280
      Width = 1455
   End
   Begin VB.ComboBox cmbAct
      Height = 315
      Left = 1440
      TabIndex = 7
      Top = 3120
      Width = 1575
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 4680
      TabIndex = 0
      TabStop = 0 'False
      ToolTipText = "Save And Exit"
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 615
      Left = 4680
      TabIndex = 4
      Top = 360
      Width = 1095
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaAPp18a.frx":0000
         Style = 1 'Graphical
         TabIndex = 6
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaAPp18a.frx":017E
         Style = 1 'Graphical
         TabIndex = 5
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 1
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
      PictureUp = "diaAPp18a.frx":0308
      PictureDn = "diaAPp18a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 2
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
      PictureUp = "diaAPp18a.frx":0594
      PictureDn = "diaAPp18a.frx":06DA
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Default Sort By Date)"
      Height = 285
      Index = 13
      Left = 3120
      TabIndex = 33
      Top = 4920
      Width = 2145
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Summary"
      Height = 285
      Index = 10
      Left = 120
      TabIndex = 31
      Top = 4560
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Sort By Vendor"
      Height = 285
      Index = 9
      Left = 120
      TabIndex = 30
      Top = 4920
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "A/P External Checks"
      Height = 285
      Index = 7
      Left = 120
      TabIndex = 28
      Top = 4320
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "A/P Computer Checks"
      Height = 285
      Index = 6
      Left = 120
      TabIndex = 27
      Top = 4080
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Starting Check #"
      Height = 285
      Index = 2
      Left = 120
      TabIndex = 24
      Top = 480
      Width = 2145
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending Check #"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 23
      Top = 840
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Start Date"
      Height = 285
      Index = 1
      Left = 120
      TabIndex = 22
      Top = 1440
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending Date"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 21
      Top = 1800
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 4
      Left = 3120
      TabIndex = 20
      Top = 1440
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 5
      Left = 3120
      TabIndex = 19
      Top = 480
      Width = 1185
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Check Vendor"
      Height = 285
      Index = 8
      Left = 120
      TabIndex = 18
      Top = 2280
      Width = 1305
   End
   Begin VB.Label lblNme
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 17
      Top = 2640
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 15
      Left = 3120
      TabIndex = 16
      Top = 2280
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Cash Account"
      Height = 285
      Index = 16
      Left = 120
      TabIndex = 15
      Top = 3120
      Width = 1305
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 14
      Top = 3480
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 17
      Left = 3120
      TabIndex = 13
      Top = 3120
      Width = 1065
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 3
      Top = 0
      Width = 2760
   End
End
Attribute VB_Name = "diaAPp18a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'***************************************************************************************
' diaAPp19a - Disbursements by GL account
'
' Notes:
'
' Created: 10/06/02 (nth)
' Revisions:
'   10/22/03 (nth) add customreport
'
'***************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***************************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo ' Cash Accounts
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaAPp18a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim rdoAct As rdoResultset
   
   On Error GoTo DiaErr1
   
   FillVendors Me
   
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1"
   bSqlRows = GetDataSet(rdoAct)
   If bSqlRows Then
      With rdoAct
         Do While Not .EOF
            AddComboStr cmbAct.hwnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbAct.ListIndex = 0
      FindAccount Me
   End If
   Set rdoAct = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbAct_click()
   FindAccount Me
End Sub

Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If Trim(cmbAct) = "" Then
      lblDsc = "Multiple Accounts Selected."
   Else
      FindAccount Me
   End If
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me
End Sub

Private Sub cmbVnd_LostFocus()
   cmbvnd = CheckLen(cmbvnd, 10)
   FindVendor Me
   If Trim(cmbvnd) = "" Then
      lblNme = "Mulitple Vendors Selected."
   Else
      FindVendor Me
   End If
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   
   sSql = ""
   optPrn.Enabled = False
   optDis.Enabled = False
   
   SetMdiReportsize MdiSect
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: '" & Secure.UserInitials & "'"
   MdiSect.crw.Formulas(2) = "StartDate='" & txtBegDte & "'"
   MdiSect.crw.Formulas(3) = "EndDate='" & txtEndDte & "'"
   MdiSect.crw.Formulas(4) = "StartCheck='" & txtBegNum & "'"
   MdiSect.crw.Formulas(5) = "EndCheck='" & txtEndNum & "'"
   
   If Trim(txtBegNum) <> "" Then _
           sSql = sSql & "{ChksTable.CHKNUMBER} >= '" & Trim(txtBegNum) & "' AND "
   If Trim(txtEndNum) <> "" Then _
           sSql = sSql & "{ChksTable.CHKNUMBER} <= '" & Trim(txtEndNum) & "' AND "
   If Trim(txtBegDte) <> "" Then _
           sSql = sSql & "{ChksTable.CHKACTUALDATE} >= #" & txtBegDte & "# AND "
   If Trim(txtEndDte) <> "" Then _
           sSql = sSql & "{ChksTable.CHKACTUALDATE} <= #" & txtEndDte & "# AND "
   
   If Trim(cmbvnd) <> "" Then
      MdiSect.crw.Formulas(12) = "CheckVendor='" & Trim(cmbvnd) & "'"
      sSql = sSql & "{ChksTable.CHKVENDOR} = '" & Compress(cmbvnd) & "' AND "
   Else
      MdiSect.crw.Formulas(12) = "CheckVendor='All'"
   End If
   
   If Trim(cmbAct) <> "" Then
      MdiSect.crw.Formulas(13) = "CheckAccount='" & Trim(cmbAct) & "'"
      sSql = sSql & "{JritTable.DCACCTNO} = '" & Compress(cmbAct) & "' AND "
   Else
      MdiSect.crw.Formulas(13) = "CheckAccount='All'"
   End If
   
   sSql = sSql & "({JrhdTable.MJTYPE} ='CC' OR {JrhdTable.MJTYPE} = 'XC')" _
          & " AND {JritTable.DCDEBIT} <> 0 AND {JritTable.DCCREDIT} = 0"
   
   MdiSect.crw.SelectionFormula = sSql
   
   If chkVnd.Value = vbChecked Then
      MdiSect.crw.Formulas(14) = "SortByDate='N'"
      MdiSect.crw.Formulas(15) = "SortByVendor='Y'"
      sCustomReport = GetCustomReport("finch06b.rpt")
   Else
      MdiSect.crw.Formulas(14) = "SortByDate='N'"
      MdiSect.crw.Formulas(15) = "SortByVendor='N'"
      sCustomReport = GetCustomReport("finch06a.rpt")
   End If
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   
   SetCrystalAction Me
   
   MouseCursor 0
   optPrn.Enabled = True
   optDis.Enabled = True
   Exit Sub
   DiaErr1:
   optPrn.Enabled = True
   optDis.Enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub
