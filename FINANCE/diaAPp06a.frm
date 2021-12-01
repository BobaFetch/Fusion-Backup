VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp06a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Average Age of Paid Invoices"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7215
   Begin VB.CheckBox optUnp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   840
      Width           =   1555
   End
   Begin VB.CheckBox optInv 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   3465
      Width           =   735
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox optChk 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox optPst 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPp06a.frx":0000
      PictureDn       =   "diaAPp06a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4320
      FormDesignWidth =   7215
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPp06a.frx":028C
      PictureDn       =   "diaAPp06a.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Invoice Date Used Otherwise)"
      Height          =   285
      Index           =   13
      Left            =   3840
      TabIndex        =   26
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   25
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   23
      Top             =   840
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   20
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unpaid Invoices?"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Invoices From:"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   1560
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Detail?"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Detail?"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting"
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   15
      Top             =   1860
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending"
      Height          =   285
      Index           =   10
      Left            =   600
      TabIndex        =   14
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Posting Date?"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "diaAPp06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaAPp06a- Average Age of Paid Invoices
'
' Created: 12/08/03 (JcW!)
' Revisions:
' 03/17/05 Fixed Qry_FillVendors (casing) in FillCombo
'
'
'*********************************************************************************
Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodVendor As Boolean
'*********************************************************************************

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(Me)
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   bGoodVendor = FindVendor(Me)
   If Not bGoodVendor Or Trim(cmbVnd) = "" Then
      cmbVnd = "ALL"
      lblNme = "***Multiple Vendors Selected***"
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      GetOptions
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   txtsDte = Format(Now, "mm/01/yy")
   txteDte = Format(Now, "mm/dd/yy")
   GetOptions
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
End Sub

Private Sub FillCombo()
   Dim RdoVed As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "Qry_FillVendors"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         Do Until .EOF
            AddComboStr cmbVnd.hwnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
      End With
   End If
   
   Set RdoVed = Nothing
   If cmbVnd.ListCount > 0 Then
      cmbVnd.ListIndex = 0
      bGoodVendor = FindVendor(Me)
   End If
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaAPp06a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sLegend As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finap06.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

   sSql = "(isnull({ChksTable.CHKNUMBER}) or {VihdTable.VIPIF} = 1 ) "
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestedBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By ESI'")
   aFormulaValue.Add CStr("'Average Age of Paid Invoices'")
   aFormulaValue.Add CStr("'Invoices From " & CStr(Trim(txtsDte) & "  Through " & Trim(txteDte)) & "'")
   
   sLegend = "Use='Use: Invoice "
   
   If optPst.Value = vbChecked Then
       aFormulaName.Add "PstInvDt"
       aFormulaValue.Add CStr("{VihdTable.VIDTRECD}")
       aFormulaName.Add "Use"
       aFormulaValue.Add CStr("'Use: Invoice " & "Posting Date '")
   Else
       aFormulaName.Add "PstInvDt"
       aFormulaValue.Add CStr("{VihdTable.VIDATE}")
       aFormulaName.Add "Use"
       aFormulaValue.Add CStr("'Use: Invoice " & "Date '")
   End If
   
   sLegend = "include: "
   
   If Trim(txtsDte) <> "" Then
      sSql = sSql & " AND {VihdTable.VIDATE} >= datetime('" & Trim(txtsDte) & "')  "
   End If
   
   If Trim(txteDte) <> "" Then
      sSql = sSql & " AND {VihdTable.VIDATE} <= Datetime('" & Trim(txteDte) & "') "
   End If
   
   If Trim(cmbVnd) <> "ALL" Then
      sSql = sSql & " AND {VihdTable.VIVENDOR} = '" & Compress(cmbVnd) & "' "
       aFormulaName.Add "Vendor"
       aFormulaValue.Add CStr("'Vendor " & CStr(cmbVnd) & "'")
   Else
       aFormulaName.Add "Vendor"
       aFormulaValue.Add CStr("'Vendor = ALL '")
   End If
   
   
   sLegend = sLegend & " Unpaid Invoices? "
   If optUnp.Value = vbUnchecked Then
      sSql = sSql & " AND {VihdTable.VIPIF} = 1 "
      sLegend = sLegend & "N"
   Else
      sLegend = sLegend & "Y"
   End If
   
   
   sLegend = sLegend & "   Invoice Detail? "
   If optInv.Value = vbUnchecked Then
       aFormulaName.Add "invDetail"
       aFormulaValue.Add "'0'"
      sLegend = sLegend & "N"
   Else
       aFormulaName.Add "invDetail"
       aFormulaValue.Add "'1'"
      sLegend = sLegend & "Y"
   End If
   
   sLegend = sLegend & "   Check Detail? "
   If optChk.Value = vbUnchecked Then
       aFormulaName.Add "chkDetail"
       aFormulaValue.Add "'0'"
      sLegend = sLegend & "N"
   Else
       aFormulaName.Add "chkDetail"
       aFormulaValue.Add "'1'"
      sLegend = sLegend & "Y"
   End If
   
   aFormulaName.Add "Title3"
   aFormulaValue.Add CStr("'" & CStr(sLegend) & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sLegend As String
   Dim sCustomReport As String
   MouseCursor 13
   On Error GoTo DiaErr1
   
   'SetMdiReportsize MdiSect
   
   sCustomReport = GetCustomReport("finap06.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "(isnull({ChksTable.CHKNUMBER}) or {VihdTable.VIPIF} = 1 ) "
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestedBy='Requested By ESI'"
   MdiSect.crw.Formulas(2) = "Title1='Average Age of Paid Invoices'"
   MdiSect.crw.Formulas(3) = "Title2='Invoices From " & Trim(txtsDte) & "  Through " & Trim(txteDte) & "'"
   
   sLegend = "Use='Use: Invoice "
   If optPst.Value = vbChecked Then
      MdiSect.crw.Formulas(4) = "PstInvDt = {VihdTable.VIDTRECD}"
      MdiSect.crw.Formulas(5) = "Use='Use: Invoice " & "Posting Date '"
   Else
      MdiSect.crw.Formulas(4) = "PstInvDt = {VihdTable.VIDATE}"
      MdiSect.crw.Formulas(5) = "Use='Use: Invoice " & "Date '"
   End If
   
   sLegend = "include: "
   
   If Trim(txtsDte) <> "" Then
      sSql = sSql & " AND {VihdTable.VIDATE} >= datetime('" & Trim(txtsDte) & "')  "
   End If
   
   If Trim(txteDte) <> "" Then
      sSql = sSql & " AND {VihdTable.VIDATE} <= Datetime('" & Trim(txteDte) & "') "
   End If
   
   If Trim(cmbVnd) <> "ALL" Then
      sSql = sSql & " AND {VihdTable.VIVENDOR} = '" & Compress(cmbVnd) & "' "
      MdiSect.crw.Formulas(6) = "Vendor ='Vendor = " & cmbVnd & "'"
   Else
      MdiSect.crw.Formulas(6) = "Vendor ='Vendor = ALL '"
   End If
   
   
   sLegend = sLegend & " Unpaid Invoices? "
   If optUnp.Value = vbUnchecked Then
      sSql = sSql & " AND {VihdTable.VIPIF} = 1 "
      sLegend = sLegend & "N"
   Else
      sLegend = sLegend & "Y"
   End If
   
   
   sLegend = sLegend & "   Invoice Detail? "
   If optInv.Value = vbUnchecked Then
      MdiSect.crw.Formulas(7) = "invDetail='0'"
      sLegend = sLegend & "N"
   Else
      MdiSect.crw.Formulas(7) = "invDetail='1'"
      sLegend = sLegend & "Y"
   End If
   
   sLegend = sLegend & "   Check Detail? "
   If optChk.Value = vbUnchecked Then
      MdiSect.crw.Formulas(8) = "chkDetail='0'"
      sLegend = sLegend & "N"
   Else
      MdiSect.crw.Formulas(8) = "chkDetail='1'"
      sLegend = sLegend & "Y"
   End If
   
   MdiSect.crw.Formulas(9) = "Title3='" & sLegend & "'"
   MdiSect.crw.SelectionFormula = sSql
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   
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


Private Sub txtEDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEdte_LostFocus()
   txteDte = CheckDate(txteDte)
End Sub

Private Sub txtSDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtSDte_LostFocus()
   txtsDte = CheckDate(txtsDte)
End Sub


Public Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   'Save by Menu Option
   sOptions = RTrim(optPst.Value) _
              & RTrim(optUnp.Value) _
              & RTrim(optInv.Value) _
              & RTrim(optChk.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optPst.Value = Val(Left(sOptions, 1))
      optUnp.Value = Val(Mid(sOptions, 2, 1))
      optInv.Value = Val(Mid(sOptions, 3, 1))
      optChk.Value = Val(Mid(sOptions, 4, 1))
   Else
      optPst.Value = vbChecked
      optUnp.Value = vbChecked
      optInv.Value = vbChecked
      optChk.Value = vbChecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
