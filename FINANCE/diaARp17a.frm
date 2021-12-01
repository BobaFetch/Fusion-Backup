VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp17a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sources of Cash"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6855
   Begin VB.Frame fraClass 
      Caption         =   "SO Classes"
      Height          =   1335
      Left            =   60
      TabIndex        =   52
      Top             =   1920
      Width           =   6735
      Begin VB.CommandButton cmdNone 
         Caption         =   "NONE"
         Height          =   315
         Left            =   3600
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ALL"
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Z"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   25
         Left            =   6135
         TabIndex        =   30
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Y"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   24
         Left            =   5625
         TabIndex        =   29
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "X"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   23
         Left            =   5130
         TabIndex        =   28
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "W"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   22
         Left            =   4635
         TabIndex        =   27
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "V"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   21
         Left            =   4140
         TabIndex        =   26
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "U"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   20
         Left            =   3660
         TabIndex        =   25
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "T"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   19
         Left            =   3165
         TabIndex        =   24
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "S"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   18
         Left            =   2670
         TabIndex        =   23
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "R"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   17
         Left            =   2175
         TabIndex        =   22
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Q"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   16
         Left            =   1665
         TabIndex        =   21
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "P"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   15
         Left            =   1170
         TabIndex        =   20
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "O"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   14
         Left            =   675
         TabIndex        =   19
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "N"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   13
         Left            =   180
         TabIndex        =   18
         Top             =   960
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "M"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   12
         Left            =   6135
         TabIndex        =   17
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "L"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   11
         Left            =   5625
         TabIndex        =   16
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "K"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   10
         Left            =   5130
         TabIndex        =   15
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "J"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   9
         Left            =   4635
         TabIndex        =   14
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "I"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   8
         Left            =   4140
         TabIndex        =   13
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "H"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   7
         Left            =   3660
         TabIndex        =   12
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "G"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   6
         Left            =   3165
         TabIndex        =   11
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "F"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   5
         Left            =   2670
         TabIndex        =   10
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "E"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   4
         Left            =   2175
         TabIndex        =   9
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "D"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   1665
         TabIndex        =   8
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "C"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   1170
         TabIndex        =   7
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "B"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   675
         TabIndex        =   6
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "A"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   435
      End
   End
   Begin VB.CheckBox chkCOSV 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   34
      Top             =   4500
      Width           =   855
   End
   Begin VB.CheckBox chkOpen 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   31
      Top             =   3420
      Width           =   855
   End
   Begin VB.CheckBox chkShipped 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   32
      Top             =   3780
      Width           =   855
   End
   Begin VB.CheckBox chkInvoiced 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   33
      Top             =   4140
      Width           =   855
   End
   Begin VB.CheckBox chkDetail 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2580
      TabIndex        =   35
      Top             =   4860
      Width           =   855
   End
   Begin VB.ComboBox cboPeriod 
      Height          =   315
      ItemData        =   "diaARp17a.frx":0000
      Left            =   2580
      List            =   "diaARp17a.frx":001C
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1500
      Width           =   1395
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      TabIndex        =   39
      Top             =   600
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaARp17a.frx":005E
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaARp17a.frx":01DC
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cboCustomer 
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
      TabIndex        =   43
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
      PictureUp       =   "diaARp17a.frx":0366
      PictureDn       =   "diaARp17a.frx":04AC
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   44
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
      PictureUp       =   "diaARp17a.frx":05F2
      PictureDn       =   "diaARp17a.frx":0738
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
      FormDesignHeight=   5295
      FormDesignWidth =   6855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "CO and SV Items?"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   4500
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open SO Items?"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   50
      Top             =   3420
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipped SO Items?"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   49
      Top             =   3780
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoiced SO Items?"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   48
      Top             =   4140
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   47
      Top             =   4860
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Size"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   46
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
      TabIndex        =   45
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2580
      TabIndex        =   42
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   41
      Top             =   420
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   40
      Top             =   1140
      Width           =   2400
   End
End
Attribute VB_Name = "diaARp17a"
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
' diaARp17a - Sources of Cash
'
' Created: 3/6/06 TEL
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

Private Sub cboCustomer_Click()
   lblName = GetCustomerName(cboCustomer)
End Sub

Private Sub cboCustomer_GotFocus()
   ComboGotFocus cboCustomer
End Sub

Private Sub cboCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
   ComboKeyUp cboCustomer, KeyCode
End Sub

Private Sub cboCustomer_LostFocus()
   lblName = GetCustomerName(cboCustomer)
End Sub

Private Sub CmdAll_Click()
   Dim i As Integer
   For i = 0 To 25
      chkClass(i).Value = vbChecked
   Next
End Sub

Private Sub cmdNone_Click()
   Dim i As Integer
   For i = 0 To 25
      chkClass(i).Value = vbUnchecked
   Next
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

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      GetOptions
      bOnLoad = False
   End If
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCombo()
   On Error GoTo DiaErr1
   
   'FillCustomers Me
   LoadComboWithCustomers cboCustomer, True
   'FindCustomer Me
   lblName = GetCustomerName(cboCustomer)
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

'Public Sub PrintReport()
'   On Error GoTo DiaErr1
'
'   MouseCursor 13
'
'   'SetMdiReportsize MdiSect
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MdiSect.crw.Formulas(2) = "Title1 = 'For " & cboPeriod & " Periods Starting " & Trim(cboStart) & "'"
'   MdiSect.crw.Formulas(3) = "Title2 = 'Customer " & Trim(cboCustomer) & ": " & lblName & "'"
'   MdiSect.crw.Formulas(4) = "Customer='" & cboCustomer & "'"
'   MdiSect.crw.Formulas(5) = "StartDate='" & cboStart & "'"
'   MdiSect.crw.Formulas(6) = "Period='" & cboPeriod & "'"
'
'   'construct string of selected SO classes
'   Dim sClasses As String
'   Dim i As Integer
'   For i = 0 To 25
'      sClasses = sClasses & chkClass(i).Value
'   Next
'   MdiSect.crw.Formulas(7) = "Classes='" & sClasses & "'"
'
'   MdiSect.crw.Formulas(8) = "ShowOpen=" & chkOpen.Value
'   MdiSect.crw.Formulas(9) = "ShowShipped=" & chkShipped.Value
'   MdiSect.crw.Formulas(10) = "ShowInvoiced=" & chkInvoiced.Value
'   MdiSect.crw.Formulas(11) = "ShowCOSV=" & chkCOSV.Value
'   MdiSect.crw.Formulas(12) = "ShowDetail=" & chkDetail.Value
'
'   sCustomReport = GetCustomReport("finar17.rpt")
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
'
'   'SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'DiaErr1:
'   sProcName = "PrintReport"
'   CurrError.Number = Err
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub

Public Sub PrintReport()
   On Error GoTo DiaErr1
   
   MouseCursor 13
   
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   optPrn.enabled = False
   optDis.enabled = False
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finar17.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "Customer"
   aFormulaName.Add "StartDate"
   aFormulaName.Add "Period"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'For " & cboPeriod & " Periods Starting " & Trim(cboStart) & "'")
   aFormulaValue.Add CStr("'Customer " & Trim(cboCustomer) & ": " & lblName & "'")
   aFormulaValue.Add CStr("'" & CStr(cboCustomer) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
   aFormulaValue.Add CStr("'" & CStr(cboPeriod) & "'")
   
   aFormulaName.Add "Classes"
   
   Dim sClasses As String
   Dim i As Integer
   Dim sSqltmp As String
   
   For i = 0 To 25
      sClasses = sClasses & chkClass(i).Value
   Next
   MdiSect.crw.Formulas(7) = "Classes='" & sClasses & "'"
   
   aFormulaValue.Add CStr("'" & CStr(sClasses) & "'")

   aFormulaName.Add "ShowOpen"
   aFormulaName.Add "ShowShipped"
   aFormulaName.Add "ShowInvoiced"
   aFormulaName.Add "ShowCOSV"
   aFormulaName.Add "ShowDetail"

   aFormulaValue.Add CStr("'" & CStr(chkOpen.Value) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkShipped.Value) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkInvoiced.Value) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkCOSV.Value) & "'")
   aFormulaValue.Add CStr("'" & CStr(chkDetail.Value) & "'")

   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'sSql = "({CustTable.CUNICKNAME} = '" & Compress(cboCustomer) & "') "
   
   If (Compress(cboCustomer) <> "<ALL>") Then
      sSqltmp = "{CustTable.CUNICKNAME} = '" & Compress(cboCustomer) & "'"
   Else
      sSqltmp = ""
   End If

   sSql = ""
   sSql = cCRViewer.GetReportSelectionFormula
   
   If (sSqltmp <> "") Then
      sSql = sSqltmp & " AND (" & sSql & ")"
   End If
   
   If (sSql <> "") Then cCRViewer.SetReportSelectionFormula sSql
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   
   'SetCrystalAction Me
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
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Customer", cboCustomer
   
   Dim sOptions As String
   sOptions = chkOpen & chkShipped & chkInvoiced & chkCOSV & chkDetail
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   
   'save so classes
   sOptions = ""
   Dim i As Integer
   For i = 0 To 25
      sOptions = sOptions & chkClass(i).Value
   Next
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Classes", sOptions
   
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", "EsiFina", Me.Name & "StartDate", defaultDate)
   cboPeriod = GetSetting("Esi2000", "EsiFina", Me.Name & "Period", "Day")
   cboCustomer = GetSetting("Esi2000", "EsiFina", Me.Name & "Customer", cboCustomer.List(0))
   lblName = GetCustomerName(cboCustomer)
   
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, "0000")
   chkOpen.Value = Mid(sOptions, 1, 1)
   chkShipped.Value = Mid(sOptions, 2, 1)
   chkInvoiced.Value = Mid(sOptions, 3, 1)
   chkCOSV.Value = Mid(sOptions, 4, 1)
   chkDetail.Value = Mid(sOptions, 5, 1)
   
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name & "Classes", "11111111111111111111111111")
   Dim i As Integer
   For i = 0 To 25
      chkClass(i).Value = Mid(sOptions, i + 1, 1)
   Next
   
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub
