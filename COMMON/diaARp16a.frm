VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARp16a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Delivery Performance Report"
   ClientHeight    =   3270
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3270
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDaysEarly 
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "0"
      Top             =   1980
      Width           =   555
   End
   Begin VB.TextBox txtDaysLate 
      Height          =   285
      Left            =   1620
      TabIndex        =   4
      Tag             =   "1"
      Text            =   "0"
      Top             =   2340
      Width           =   555
   End
   Begin VB.ComboBox cboCustomer 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Tag             =   "3"
      Top             =   360
      Width           =   1555
   End
   Begin VB.CheckBox optDtl 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1620
      TabIndex        =   5
      Top             =   2700
      Width           =   855
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1140
      Width           =   1215
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3270
      FormDesignWidth =   5835
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4680
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp16a.frx":0000
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
         Picture         =   "diaARp16a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
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
      PictureUp       =   "diaARp16a.frx":0308
      PictureDn       =   "diaARp16a.frx":044E
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
      PictureUp       =   "diaARp16a.frx":0594
      PictureDn       =   "diaARp16a.frx":06DA
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   380
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Early"
      Height          =   285
      Index           =   5
      Left            =   2340
      TabIndex        =   21
      Top             =   1980
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   1980
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Late"
      Height          =   285
      Index           =   3
      Left            =   2340
      TabIndex        =   19
      Top             =   2340
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   2340
      Width           =   1005
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1620
      TabIndex        =   17
      Top             =   795
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   435
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2700
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1140
      Width           =   825
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
Attribute VB_Name = "diaARp16a"
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
' diaARp16a - Customer Delivery Performance
'
' Notes:
'
' Created: 01/29/03 (nth)
' Revisions:
'   03/04/04 (nth) Added detail for Linda
'   03/05/04 (nth) Added cash basis version of report for Linda
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboCustomer_Click()
   If cboCustomer = "ALL" Then lblName = "All Customers" Else lblName = GetCustomerName(cboCustomer)
End Sub



Private Sub cboCustomer_GotFocus()
   ComboGotFocus cboCustomer
End Sub

Private Sub cboCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
   'ComboKeyUp cboCustomer, KeyCode 'BBS removed this on 03/25/2010 for Ticket #27401
   'lblName = GetCustomerName(cboCustomer) 'BBS removed this on 03/25/2010 for Ticket #27401
End Sub

Private Sub cboCustomer_LostFocus()
    If cboCustomer = "" Then cboCustomer = "ALL"
    If cboCustomer = "ALL" Then lblName = "All Customers" Else lblName = GetCustomerName(cboCustomer)
End Sub

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
     LoadComboWithCustomers cboCustomer, False
     GetOptions
      GetOptions
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   '    cboStart = Format(Now, "mm/01/yy")
   '    cboEnd = Format(Now, "mm/31/yy")
   '    GetOptions
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARp16a = Nothing
End Sub

'Private Sub optCsh_Click()
'    If optCsh Then
'        optDtl = vbChecked
'        optDtl.enabled = False
'    Else
'        optDtl.enabled = True
'    End If
'End Sub
'

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
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   optPrn.Enabled = False
   optDis.Enabled = False
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finar16.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   If Compress(cboCustomer) = "" Then cboCustomer = "ALL"
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Customer"
   aFormulaName.Add "StartDate"
   aFormulaName.Add "EndDate"
   aFormulaName.Add "ShowDetail"
   aFormulaName.Add "DaysEarly"
   aFormulaName.Add "DaysLate"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer Delivery Performance'")
   aFormulaValue.Add CStr("'Customer " & CStr(cboCustomer & ": " & lblName) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboCustomer) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboEnd) & "'")
   aFormulaValue.Add optDtl.Value
   aFormulaValue.Add txtDaysEarly
   aFormulaValue.Add txtDaysLate
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   optPrn.Enabled = True
   optDis.Enabled = True
   sSql = "{SoitTable.ITPSSHIPPED} = 1.00 and {SoitTable.ITACTUAL} in CDateTime ({@StartDate}) to CDateTime ({@EndDate})"
   If Compress(cboCustomer) <> "ALL" Then sSql = sSql & " AND {SohdTable.SOCUST} = '" & cboCustomer & "'" 'BBS Added for Ticket #37097
    
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
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub SaveOptions()
   Dim strRegApp As String
   strRegApp = GetRegistryAppTitle()
   
   SaveSetting "Esi2000", strRegApp, Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", strRegApp, Me.Name & "EndDate", cboEnd
   SaveSetting "Esi2000", strRegApp, Me.Name & "Customer", cboCustomer
   SaveSetting "Esi2000", strRegApp, Me.Name & "DaysEarly", txtDaysEarly
   SaveSetting "Esi2000", strRegApp, Me.Name & "DaysLate", txtDaysLate
   
   Dim sOptions As String
   sOptions = optDtl
   SaveSetting "Esi2000", strRegApp, Me.Name, sOptions
   SaveSetting "Esi2000", strRegApp, Me.Name & "_Printer", lblPrinter
End Sub

Private Sub GetOptions()
   Dim defaultDate As String
   Dim strRegApp As String
   strRegApp = GetRegistryAppTitle()
   
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", strRegApp, Me.Name & "StartDate", defaultDate)
   cboEnd = GetSetting("Esi2000", strRegApp, Me.Name & "EndDate", defaultDate)
   cboCustomer = GetSetting("Esi2000", strRegApp, Me.Name & "Customer", cboCustomer.List(0))
   If cboCustomer = "ALL" Then lblName = "All Customers" Else lblName = GetCustomerName(cboCustomer)
   txtDaysEarly = GetSetting("Esi2000", strRegApp, Me.Name & "DaysEarly", 0)
   txtDaysLate = GetSetting("Esi2000", strRegApp, Me.Name & "DaysLate", 0)
   
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", strRegApp, Me.Name, sOptions)
   If Len(Trim(sOptions)) >= 2 Then
      optDtl.Value = Mid(sOptions, 2, 1)
   Else
      optDtl.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", strRegApp, Me.Name & "_Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub cboEnd_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cboEnd_LostFocus()
   cboEnd = CheckDateEx(cboEnd)
End Sub

Private Sub cboStart_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cboStart_LostFocus()
   cboStart = CheckDateEx(cboStart)
End Sub

