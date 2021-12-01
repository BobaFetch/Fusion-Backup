VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CustSh01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print/Display Schedules"
   ClientHeight    =   3210
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3210
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "CustSh01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CustSh01b.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optVal 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optPod 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   2400
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optSod 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1440
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CustSh01b.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "CustSh01b.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3210
      FormDesignWidth =   7215
   End
   Begin VB.Label Label2 
      Caption         =   "through"
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label txtBeg 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Value"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing"
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Detail"
      Height          =   285
      Index           =   6
      Left            =   360
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label txtEnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4440
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label cmbWcn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Complete From"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   7
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "CustSh01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




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



Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move CustSh01.Left + 800, CustSh01.Top + 1000
   FormatControls
   GetOptions
   On Error Resume Next
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   CustSh01.optReport.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set CustSh01b = Nothing
   
End Sub



Private Sub PrintReport()
   Dim sPrices As String
   MouseCursor 13
   On Error GoTo DiaErr1
   
   
   If optVal.Value = vbChecked Then sPrices = "Y" Else sPrices = ""
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   sCustomReport = GetCustomReport("custsh01.rpt")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "Option1"
    aFormulaName.Add "ShowSod"
    aFormulaName.Add "ShowPod"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & cmbWcn & ", Through " & txtEnd & "'")
    aFormulaValue.Add CStr("'" & sPrices & "'")
    aFormulaValue.Add optSod.Value
    aFormulaValue.Add optPod.Value
    
    
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
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
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'Private Sub PrintReport()
'   Dim sPrices As String
'   Dim sWindows As String
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   SetMdiReportsize MDISect
'   ' sWindows = GetWindowsDir()
'   sWindows = "c:\windows\"
'
'   If optVal.Value = vbChecked Then sPrices = "Y" Else sPrices = ""
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Includes='" & cmbWcn & ", Through " _
'                        & txtEnd & "'"
'   MDISect.Crw.Formulas(2) = "Option1='" & sPrices & "'"
'
'   'V Creates an error if left unblocked 4/5/06
'   MDISect.Crw.ReportFileName = sReportPath & "custsh01.rpt"
'   If optSod.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.2;F;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.1.3;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.2;T;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.1.3;T;;;"
'   End If
'   If optPod.Value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(3) = "GROUPHDR.1.4;F;;;"
'      MDISect.Crw.SectionFormat(4) = "DETAIL.0.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(3) = "GROUPHDR.1.4;T;;;"
'      MDISect.Crw.SectionFormat(4) = "DETAIL.0.0;T;;;"
'   End If
'
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optSod.Value)) & Trim(str(optPod.Value))
   SaveSetting "Esi2000", "EsiProd", "custsh01b", sOptions
   SaveSetting "Esi2000", "EsiProd", "custsh01bprn", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "custsh01b", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optSod.Value = Val(Left(sOptions, 1))
      optPod.Value = Val(Right(sOptions, 1))
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "custsh01bprn", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub
