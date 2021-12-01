VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARp15a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "B & O Tax Liability"
   ClientHeight    =   2490
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2490
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSte 
      Height          =   315
      ItemData        =   "diaARp15a.frx":0000
      Left            =   1320
      List            =   "diaARp15a.frx":0002
      TabIndex        =   0
      Top             =   540
      Width           =   855
   End
   Begin VB.CheckBox optDtl 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1980
      Width           =   855
   End
   Begin VB.CheckBox optCsh 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1620
      Width           =   855
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1260
      Width           =   1095
   End
   Begin VB.ComboBox txtStart 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "4"
      Top             =   900
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4860
      Top             =   1620
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2490
      FormDesignWidth =   5835
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp15a.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaARp15a.frx":0182
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   9
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
      PictureUp       =   "diaARp15a.frx":030C
      PictureDn       =   "diaARp15a.frx":0452
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   10
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
      PictureUp       =   "diaARp15a.frx":0598
      PictureDn       =   "diaARp15a.frx":06DE
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "State Code"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Otherwise Use Accrual) "
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   15
      Top             =   1620
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Basis"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   825
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaARp15a"
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
' diaARp15a - B&O Tax Liability (Report)
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

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FillStates Me
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   txtstart = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/31/yy")
   GetOptions
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
   Set diaARp15a = Nothing
End Sub

Private Sub optCsh_Click()
   If optCsh Then
      optDtl = vbChecked
      optDtl.enabled = False
   Else
      optDtl.enabled = True
   End If
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
   If optCsh Then
      sCustomReport = GetCustomReport("finar15b.rpt")
   Else
      sCustomReport = GetCustomReport("finar15a.rpt")
   End If
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title2"
   aFormulaName.Add "StartDate"
   aFormulaName.Add "EndDate"
   aFormulaName.Add "ShowDetail"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'For " & CStr(cmbSte.Text & " " & txtstart & " Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'" & txtstart & "'")
   aFormulaValue.Add CStr("'" & txtEnd & "'")
   aFormulaValue.Add optDtl.Value
   
   'MdiSect.crw.Formulas(5) = "CutOff='" & txtEnd & "'"
   
   sSql = ""
   If (Trim(cmbSte.Text) <> "") Then
      sSql = "{TxcdTable.TAXSTATE} = '" & cmbSte.Text & "' AND "
   End If
   
   If optCsh Then
      sSql = sSql & "{CashTable.CARCDATE} >= #" & txtstart & "# AND " _
             & "{CashTable.CARCDATE} <= #" & txtEnd & "#  AND {CihdTable.INVCANCELED} = 0"
      aFormulaName.Add "CutOff"
      aFormulaValue.Add CStr("'" & txtEnd & "'")
   Else
      sSql = sSql & "{CihdTable.INVDATE} >= #" & txtstart & "# AND " _
             & "{CihdTable.INVDATE} <= #" & txtEnd & "#   AND {CihdTable.INVCANCELED} = 0"
   End If
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
      
   optPrn.enabled = True
   optDis.enabled = True
   
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

Private Sub PrintReport1()
   MouseCursor 13
   On Error GoTo DiaErr1
   optPrn.enabled = False
   optDis.enabled = False
   'SetMdiReportsize MdiSect
   If optCsh Then
      sCustomReport = GetCustomReport("finar15b.rpt")
   Else
      sCustomReport = GetCustomReport("finar15a.rpt")
   End If
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   'MdiSect.crw.Formulas(2) = "Title1='Accrual Basis B & O Tax Liablility Report'"
   MdiSect.crw.Formulas(3) = "Title2='For " & cmbSte.Text & " " & txtstart & " Through " & txtEnd & "'"
   MdiSect.crw.Formulas(4) = "ShowDetail=" & optDtl.Value
   If optCsh Then
      sSql = "{TxcdTable.TAXSTATE} = '" & cmbSte.Text & "' AND " _
             & "{CashTable.CARCDATE} >= #" & txtstart & "# AND " _
             & "{CashTable.CARCDATE} <= #" & txtEnd & "#"
      MdiSect.crw.Formulas(5) = "CutOff='" & txtEnd & "'"
   Else
      sSql = "{TxcdTable.TAXSTATE} = '" & cmbSte.Text & "' AND " _
             & "{CihdTable.INVDATE} >= #" & txtstart & "# AND " _
             & "{CihdTable.INVDATE} <= #" & txtEnd & "#"
   End If
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   optPrn.enabled = True
   optDis.enabled = True
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = optCsh & optDtl
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) >= 2 Then
      optCsh.Value = Mid(sOptions, 1, 1)
      optDtl.Value = Mid(sOptions, 2, 1)
   Else
      optCsh.Value = vbUnchecked
      optDtl.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtstart_LostFocus()
   txtstart = CheckDate(txtstart)
End Sub
