VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp25a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uses of Cash"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6855
   Begin VB.CheckBox chkOpenPOs 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox chkReceivedPOs 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.CheckBox chkInvoices 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox chkDetail 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.ComboBox cboPeriod 
      Height          =   315
      ItemData        =   "diaAPp25a.frx":0000
      Left            =   2580
      List            =   "diaAPp25a.frx":001C
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1500
      Width           =   1395
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   600
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaAPp25a.frx":005E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaAPp25a.frx":01DC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cboVendor 
      Height          =   315
      Left            =   2580
      TabIndex        =   0
      Tag             =   "3"
      Top             =   345
      Width           =   1555
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   2580
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1395
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   14
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
      PictureUp       =   "diaAPp25a.frx":0366
      PictureDn       =   "diaAPp25a.frx":04AC
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   15
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
      PictureUp       =   "diaAPp25a.frx":05F2
      PictureDn       =   "diaAPp25a.frx":0738
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4740
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3420
      FormDesignWidth =   6855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open PO Items?"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Received PO Items?"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices and Invoiced PO Items?"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Size"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2580
      TabIndex        =   13
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Nickname"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   420
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1140
      Width           =   2400
   End
End
Attribute VB_Name = "diaAPp25a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' diaAPp24a - Uses of Cash
'
' Created: 3/5/06 TEL
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim sMsg As String
Public bRemote As Byte

' Key Handeling
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cboVendor_Click()
   lblName = GetVendorName(cboVendor)
End Sub

Private Sub cboVendor_GotFocus()
   ComboGotFocus cboVendor
End Sub

Private Sub cboVendor_KeyUp(KeyCode As Integer, Shift As Integer)
   ComboKeyUp cboVendor, KeyCode
End Sub

Private Sub cboVendor_LostFocus()
   lblName = GetVendorName(cboVendor)
End Sub

Private Sub Form_Load()
   If bRemote Then
      Me.WindowState = vbMinimized
   Else
      FormLoad Me
      FormatControls
      bOnLoad = True
   End If
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      GetOptions
      bOnLoad = False
   End If
End Sub

Public Sub FillCombo()
   On Error GoTo DiaErr1
   
   'FillVendors Me
   LoadComboWithVendors cboVendor, True
   'FindVendor Me
   lblName = GetVendorName(cboVendor)
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
   On Error Resume Next
   SaveOptions
   If Not bRemote Then
      FormUnload
   End If
   bRemote = False
   Set diaARp09a = Nothing
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub optDis_Click()
   If Not bRemote Then
      PrintReport
   End If
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboStart_GotFocus()
   SelectFormat Me
End Sub

Private Sub cboStart_LostFocus()
   If Trim(cboStart) <> "" Then
      cboStart = CheckDate(cboStart)
   End If
End Sub

Public Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   On Error GoTo DiaErr1
   If Trim(cboVendor) = "" Then
      sMsg = "Please Select A Vendor."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "Vendor"
    aFormulaName.Add "StartDate"
    aFormulaName.Add "Period"
    aFormulaName.Add "ShowOpen"
    aFormulaName.Add "ShowReceived"
    aFormulaName.Add "ShowInvoiced"
    aFormulaName.Add "ShowDetail"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'For " & CStr(cboPeriod & " Periods Starting " & Trim(cboStart)) & "'")
    aFormulaValue.Add CStr("'Vendor " & CStr(Trim(cboVendor) & ": " & lblName) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboVendor) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboPeriod) & "'")
    aFormulaValue.Add chkOpenPOs.Value
    aFormulaValue.Add chkReceivedPOs.Value
    aFormulaValue.Add chkInvoices.Value
    aFormulaValue.Add chkDetail.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finap15.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub PrintReport1()
   On Error GoTo DiaErr1
   If Trim(cboVendor) = "" Then
      sMsg = "Please Select A Vendor."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   
   'setmdireportsizemdisect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1 = 'For " & cboPeriod & " Periods Starting " & Trim(cboStart) & "'"
   MdiSect.crw.Formulas(3) = "Title2 = 'Vendor " & Trim(cboVendor) & ": " & lblName & "'"
   MdiSect.crw.Formulas(4) = "Vendor='" & cboVendor & "'"
   MdiSect.crw.Formulas(5) = "StartDate='" & cboStart & "'"
   MdiSect.crw.Formulas(6) = "Period='" & cboPeriod & "'"
   MdiSect.crw.Formulas(7) = "ShowOpen=" & chkOpenPOs.Value
   MdiSect.crw.Formulas(8) = "ShowReceived=" & chkReceivedPOs.Value
   MdiSect.crw.Formulas(9) = "ShowInvoiced=" & chkInvoices.Value
   MdiSect.crw.Formulas(19) = "ShowDetail=" & chkDetail.Value
   
   sCustomReport = GetCustomReport("finap15.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   'setcrystalaction me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Period", cboPeriod
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Vendor", cboVendor
   
   Dim sOptions As String
   sOptions = chkOpenPOs & chkReceivedPOs & chkInvoices & chkDetail
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", "EsiFina", Me.Name & "StartDate", defaultDate)
   cboPeriod = GetSetting("Esi2000", "EsiFina", Me.Name & "Period", "Day")
   cboVendor = GetSetting("Esi2000", "EsiFina", Me.Name & "Vendor", cboVendor.List(0))
   lblName = GetVendorName(cboVendor)
   
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, "0000")
   chkOpenPOs.Value = Mid(sOptions, 1, 1)
   chkReceivedPOs.Value = Mid(sOptions, 2, 1)
   chkInvoices.Value = Mid(sOptions, 3, 1)
   chkDetail.Value = Mid(sOptions, 4, 1)
   
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub
