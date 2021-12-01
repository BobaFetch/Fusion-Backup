VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QuickBooks ® Customer Equivalents (Report)"
   ClientHeight    =   2775
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optES 
      Caption         =   "ES/2002 Customers"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton OptQB 
      Caption         =   "QuickBooks Customers"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   2175
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2775
      FormDesignWidth =   6810
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Lending Character  Search"
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp12a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaARp12a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   7
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
      PictureUp       =   "diaARp12a.frx":0308
      PictureDn       =   "diaARp12a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   8
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
      PictureUp       =   "diaARp12a.frx":0594
      PictureDn       =   "diaARp12a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   405
      Index           =   0
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer That Begin With"
      Height          =   405
      Index           =   5
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1305
   End
End
Attribute VB_Name = "diaARp12a"
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
' diaARp12a - QuickBooks Customer Equivilents
'
' Created: 6/19/02 (nth)
' Revisions:
'
'
'*********************************************************************************

Dim bOnLoad As Byte

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
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARp12a = Nothing
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
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
  
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
  
   If optES Then
       sCustomReport = GetCustomReport("finar12a.rpt")
   Else
       sCustomReport = GetCustomReport("finar12b.rpt")
   End If
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   If Len(Trim(txtCst)) Then
      sSql = "{CustTable.CUNAME} LIKE '" & Trim(txtCst) & "*'"
      cCRViewer.SetReportSelectionFormula sSql
   End If
   
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
   
   If optES Then
      MdiSect.crw.ReportFileName = sReportPath & "finar12a.rpt"
   Else
      MdiSect.crw.ReportFileName = sReportPath & "finar12b.rpt"
   End If
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   
   If Len(Trim(txtCst)) Then
      sSql = "{CustTable.CUNAME} LIKE '" & Trim(txtCst) & "*'"
      MdiSect.crw.SelectionFormula = sSql
   End If
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
