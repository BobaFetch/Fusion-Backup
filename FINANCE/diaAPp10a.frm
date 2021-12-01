VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPp10a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchases By GL Account (Report) "
   ClientHeight    =   4260
   ClientLeft      =   2115
   ClientTop       =   435
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4260
   ScaleWidth      =   6810
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Tag             =   "2"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox chkPO 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   3615
      Width           =   975
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox chkPst 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "4"
      ToolTipText     =   "Contains Customers With Aging"
      Top             =   600
      Width           =   1080
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   8
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaAPp10a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaAPp10a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
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
      PictureUp       =   "diaAPp10a.frx":0308
      PictureDn       =   "diaAPp10a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3720
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4260
      FormDesignWidth =   6810
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   11
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
      PictureUp       =   "diaAPp10a.frx":0594
      PictureDn       =   "diaAPp10a.frx":06DA
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   20
      Top             =   2400
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include P.O. Number?"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   3720
      TabIndex        =   18
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use The Posting Date?"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(If Not, The Invoice Date Is Used)"
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   14
      Top             =   3360
      Width           =   2505
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Purchases From"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1905
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaAPp10a"
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
' diaAPp10a - Purchases By GL Account
'
' Notes:
'
' Created: 10/10/03 (JCW)
' Revisions:
'   10/22/03 (nth) Added lblDsc and lblNme
'   10/22/03 (nth) Added custom report
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbAct_Click()
   FindAccount Me
End Sub

Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If Trim(cmbAct) = "" Then
      lblDsc = "Muliple Accounts Selected."
   Else
      FindAccount Me
   End If
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me, True
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Trim(cmbVnd) = "" Then
      lblNme = "Multiple Vendors Selected."
   Else
      FindVendor Me, True
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   sCurrForm = Caption
   bOnLoad = True
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaAPp10a = Nothing
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

Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo DiaErr1
'   SetMdiReportsize MdiSect
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim aSortList As New Collection
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "BegDate"
   aFormulaName.Add "EndDate"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'Purchases By GL Account Report'")
   aFormulaValue.Add CStr("'From" & CStr(txtBeg & " Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(txtBeg)) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(txtEnd)) & "'")

'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "Title1='Purchases By GL Account Report'"
'   MdiSect.Crw.Formulas(3) = "Title2='From " & txtBeg & " Through " & txtEnd & "'"
'   MdiSect.Crw.Formulas(4) = "BegDate='" & Trim(txtBeg) & "'"
'   MdiSect.Crw.Formulas(5) = "EndDate='" & Trim(txtEnd) & "'"
'   MdiSect.Crw.ReportFileName = sReportPath & "finap10a.rpt"
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finap10a.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "({JritTable.DCDEBIT} <> 0 OR {JritTable.DCCREDIT} <> 0)  and {JrhdTable.MJTYPE} = 'pj'"
   If Trim(cmbVnd) <> "" Then
'      MdiSect.Crw.Formulas(6) = "Vendor='Vendor: " & cmbVnd & "'"
'      MdiSect.Crw.Formulas(7) = "VendorRef ='" & Compress(cmbVnd) & "'"
      aFormulaName.Add "Vendor"
      aFormulaName.Add "VendorRef"
      aFormulaValue.Add CStr("'Vendor: " & CStr(cmbVnd) & "'")
      aFormulaValue.Add CStr("'" & CStr(Compress(cmbVnd)) & "'")
      sSql = sSql & " and {VndrTable.VEREF} = '" & Compress(cmbVnd) & "' "
   Else
'      MdiSect.Crw.Formulas(6) = "Vendor='Vendor: All'"
'      MdiSect.Crw.Formulas(7) = "VendorRef='All'"
      aFormulaName.Add "Vendor"
      aFormulaName.Add "VendorRef"
      aFormulaValue.Add CStr("'Vendor: All'")
      aFormulaValue.Add CStr("'All'")
   End If
   If Trim(cmbAct) <> "" Then
      sSql = sSql & " and {GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "' "
'      MdiSect.Crw.Formulas(8) = "incAccts='Account: " & Trim(cmbAct) & "'"
'      MdiSect.Crw.Formulas(9) = "account='" & Trim(cmbAct) & "'"
      aFormulaName.Add "incAccts"
      aFormulaName.Add "account"
      aFormulaValue.Add CStr("'Account:  " & CStr(Trim(cmbAct)) & "'")
      aFormulaValue.Add CStr("'" & CStr(Trim(cmbAct)) & "'")
   Else
'      MdiSect.Crw.Formulas(8) = "incAccts='Accounts: All'"
'      MdiSect.Crw.Formulas(9) = "account='All'"
      aFormulaName.Add "incAccts"
      aFormulaName.Add "account"
      aFormulaValue.Add CStr("'Accounts: All'")
      aFormulaValue.Add CStr("'All'")
   End If
   If chkPst.Value = vbUnchecked Then
      MdiSect.crw.SortFields(0) = "+{VihdTable.VIDATE}"
      sSql = sSql & " and {VihdTable.VIDATE} >= DateTime ('" & Trim(txtBeg) & "') and " _
             & "{VihdTable.VIDATE} <= DateTime ('" & Trim(txtEnd) & "')"
'      MdiSect.Crw.Formulas(10) = "date = 'Inv Date'"
        aFormulaName.Add "date"
        aFormulaValue.Add CStr("'Inv Date'")
   Else
'      MdiSect.crw.SortFields(0) = "+{VihdTable.VIDTRECD}"
      MdiSect.crw.SortFields(0) = "+{JritTable.DCDATE}"
'      sSql = sSql & " and {VihdTable.VIDTRECD} >= DateTime ('" & Trim(txtBeg) & "') and " _
'             & "{VihdTable.VIDTRECD} <= DateTime ('" & Trim(txtEnd) & "')"
      sSql = sSql & " and {JritTable.DCDATE} >= DateTime ('" & Trim(txtBeg) & "') and " _
             & "{JritTable.DCDATE} <= DateTime ('" & Trim(txtEnd) & "')"
'      MdiSect.Crw.Formulas(10) = "date = 'Post Dt.'"
        aFormulaName.Add "date"
        aFormulaValue.Add CStr("'Post Dt.'")
   End If
   If chkPO.Value = vbChecked Then
'      MdiSect.Crw.Formulas(11) = "PO = '1'"
        aFormulaName.Add "PO"
        aFormulaValue.Add "'1'"
   Else
'      MdiSect.Crw.Formulas(11) = "PO = '0'"
        aFormulaName.Add "date"
        aFormulaValue.Add "'0'"
   End If

'   MdiSect.Crw.SortFields(1) = "+{JritTable.DCPONUMBER}"
'   MdiSect.Crw.SortFields(2) = "+{JritTable.DCPORELEASE}"
'   MdiSect.Crw.SortFields(3) = "+{JritTable.DCPOITEM}"
'   MdiSect.Crw.SortFields(0) = "+{JritTable.DCPOITREV}"
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   aSortList.Add "DCPOITREV"
   aSortList.Add "DCPONUMBER"
   aSortList.Add "DCPORELEASE"
   aSortList.Add "DCPOITEM"
   cCRViewer.SetSortFields aSortList
   
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
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='Purchases By GL Account Report'"
   MdiSect.crw.Formulas(3) = "Title2='From " & txtBeg & " Through " & txtEnd & "'"
   MdiSect.crw.Formulas(4) = "BegDate='" & Trim(txtBeg) & "'"
   MdiSect.crw.Formulas(5) = "EndDate='" & Trim(txtEnd) & "'"
   MdiSect.crw.ReportFileName = sReportPath & "finap10a.rpt"
   sSql = "{JritTable.DCDEBIT} <> 0 and {JrhdTable.MJTYPE} = 'pj'"
   If Trim(cmbVnd) <> "" Then
      MdiSect.crw.Formulas(6) = "Vendor='Vendor: " & cmbVnd & "'"
      MdiSect.crw.Formulas(7) = "VendorRef ='" & Compress(cmbVnd) & "'"
      sSql = sSql & " and {VndrTable.VEREF} = '" & Compress(cmbVnd) & "' "
   Else
      MdiSect.crw.Formulas(6) = "Vendor='Vendor: All'"
      MdiSect.crw.Formulas(7) = "VendorRef='All'"
   End If
   If Trim(cmbAct) <> "" Then
      sSql = sSql & " and {GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "' "
      MdiSect.crw.Formulas(8) = "incAccts='Account: " & Trim(cmbAct) & "'"
      MdiSect.crw.Formulas(9) = "account='" & Trim(cmbAct) & "'"
   Else
      MdiSect.crw.Formulas(8) = "incAccts='Accounts: All'"
      MdiSect.crw.Formulas(9) = "account='All'"
   End If
   If chkPst.Value = vbUnchecked Then
      MdiSect.crw.SortFields(0) = "+{VihdTable.VIDATE}"
      sSql = sSql & " and {VihdTable.VIDATE} >= DateTime ('" & Trim(txtBeg) & "') and " _
             & "{VihdTable.VIDATE} <= DateTime ('" & Trim(txtEnd) & "')"
      MdiSect.crw.Formulas(10) = "date = 'Inv Date'"
   Else
      MdiSect.crw.SortFields(0) = "+{VihdTable.VIDTRECD}"
      sSql = sSql & " and {VihdTable.VIDTRECD} >= DateTime ('" & Trim(txtBeg) & "') and " _
             & "{VihdTable.VIDTRECD} <= DateTime ('" & Trim(txtEnd) & "')"
      MdiSect.crw.Formulas(10) = "date = 'Post Dt.'"
   End If
   If chkPO.Value = vbChecked Then
      MdiSect.crw.Formulas(11) = "PO = '1'"
   Else
      MdiSect.crw.Formulas(11) = "PO = '0'"
   End If
   MdiSect.crw.SortFields(1) = "+{JritTable.DCPONUMBER}"
   MdiSect.crw.SortFields(2) = "+{JritTable.DCPORELEASE}"
   MdiSect.crw.SortFields(3) = "+{JritTable.DCPOITEM}"
   MdiSect.crw.SortFields(0) = "+{JritTable.DCPOITREV}"
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillCombo()
   Dim rdoAct As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccountCombo"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         Do While Not .EOF
            AddComboStr cmbAct.hWnd, Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
      cmbAct.ListIndex = 0
      FindAccount Me
   End If
   Set rdoAct = Nothing
   FillVendors Me
   Exit Sub
   
DiaErr1:
   sProcName = "getinvacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(chkPst.Value) _
              & RTrim(chkPO.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      chkPst.Value = Val(Mid(sOptions, 1, 1))
      chkPO.Value = Val(Mid(sOptions, 2, 1))
   Else
      chkPst.Value = vbChecked
      chkPO.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
