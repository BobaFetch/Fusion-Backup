VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp11a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Tax Liabilty"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6615
   Begin VB.CheckBox optDet 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox cmbSte 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CheckBox optCsh 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "4"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   480
      Width           =   1335
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaARp11a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaARp11a.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3105
      FormDesignWidth =   6615
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   15
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
      PictureUp       =   "diaARp11a.frx":0308
      PictureDn       =   "diaARp11a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   16
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
      PictureUp       =   "diaARp11a.frx":0594
      PictureDn       =   "diaARp11a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   18
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Audit Detail?"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2085
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "For What State?"
      Height          =   315
      Index           =   21
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Otherwise Accrual Basis)"
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   12
      Top             =   1860
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Cash Basis?"
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1845
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period From"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "diaARp11a"
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

'**************************************************************************************
' Sales Tax Liability Report
'
' Created: 01/27/04 (JCW)
' Revisions:
'
'**************************************************************************************

Dim bOnLoad As Byte


'**************************************************************************************

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbSte_LostFocus()
   cmbSte = CheckLen(cmbSte, 2)
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub


Private Sub FillCombo()
   On Error GoTo DiaErr1
   FillStates Me
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
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
   sCurrForm = Caption
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
   bOnLoad = True
   GetOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaARp11a = Nothing
End Sub

Private Sub PrintReport()
   Dim sBasis As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If optCsh.Value = 1 Then
      sBasis = "Cash Basis "
   Else
      sBasis = "Accrual Basis "
   End If
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "Det"
   aFormulaName.Add "Csh"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(sBasis) & "Sales Tax Liability'")
   aFormulaValue.Add CStr("'Period From " & CStr(txtBeg & "  Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'" & CStr(optDet.Value) & "'")
   aFormulaValue.Add CStr("'" & CStr(optCsh.Value) & "'")
   
   'CUSTOM REPORT
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finar11a")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   If Trim(sCustomReport) <> "" Then
       cCRViewer.SetReportFileName sCustomReport, sReportPath
       cCRViewer.SetReportTitle = sCustomReport
   Else
       cCRViewer.SetReportFileName sCustomReport, sReportPath
       cCRViewer.SetReportTitle = sCustomReport
   End If
   
   sSql = "{JritTable.DCDEBIT} <> 0 "
   
   sSql = sSql & " and {@Date} >= datetime('" & Trim(txtBeg) & "') and {@Date} <= datetime('" _
          & Trim(txtEnd) & "')"
   
   If Trim(cmbSte) <> "" Then
      sSql = sSql & " and {TxcdTable.TAXSTATE} = '" & Trim(cmbSte) & "'"
   End If
   
   If optCsh.Value = 1 Then
      sSql = sSql & " and {CihdTable.INVPIF} = 1 "
   End If
   
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
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   Dim sBasis As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   If optCsh.Value = 1 Then
      sBasis = "Cash Basis "
   Else
      sBasis = "Accrual Basis "
   End If
   
   
   'SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestedBy='Requested By:" & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='" & sBasis & "Sales Tax Liability'"
   MdiSect.crw.Formulas(3) = "Title2='Period From " & txtBeg & "  Through " & txtEnd & "'"
   MdiSect.crw.Formulas(4) = "Det='" & CStr(optDet.Value) & "'"
   MdiSect.crw.Formulas(5) = "Csh='" & CStr(optCsh.Value) & "'"
   
   'CUSTOM REPORT
   sCustomReport = GetCustomReport("finar11a")
   
   If Trim(sCustomReport) <> "" Then
      MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   Else
      MdiSect.crw.ReportFileName = sReportPath & "finar11a.rpt"
   End If
   
   sSql = "{JritTable.DCDEBIT} <> 0 "
   
   sSql = sSql & " and {@Date} >= datetime('" & Trim(txtBeg) & "') and {@Date} <= datetime('" _
          & Trim(txtEnd) & "')"
   
   If Trim(cmbSte) <> "" Then
      sSql = sSql & " and {TxcdTable.TAXSTATE} = '" & Trim(cmbSte) & "'"
   End If
   
   If optCsh.Value = 1 Then
      sSql = sSql & " and {CihdTable.INVPIF} = 1 "
   End If
   
   
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
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

Private Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   'Save by Menu Option
   sOptions = RTrim(optCsh.Value) _
              & RTrim(optDet.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optCsh.Value = Val(Mid(sOptions, 1, 1))
      optDet.Value = Val(Mid(sOptions, 2, 1))
   Else
      optCsh.Value = vbUnchecked
      optDet.Value = vbChecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub
