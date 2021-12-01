VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routings By Routing Number"
   ClientHeight    =   2865
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1080
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "RoutRTp02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "RoutRTp02a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   645
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   645
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2865
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5760
      TabIndex        =   11
      Top             =   1080
      Width           =   1404
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing Comments?"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing Operations?"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1110
      Width           =   1815
   End
End
Attribute VB_Name = "RoutRTp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/13/05 Corrected query criteria
Option Explicit

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbRte_LostFocus()
   cmbRte = CheckLen(cmbRte, 30)
   If Len(cmbRte) = 0 Then cmbRte = "ALL"
   
End Sub


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
   MouseCursor 0
   If bOnLoad = 1 Then FillRoutings
   cmbRte = "ALL"
   bOnLoad = 0
   MDISect.lblBotPanel = Caption
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set RoutRTp02a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sRout As String
   Dim sWindows As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
'   SetMdiReportsize MDISect
   sRout = Compress(cmbRte)
   If Len(sRout) = 0 Then
      cmbRte = "ALL"
      sRout = ""
   Else
      If sRout = "ALL" Then sRout = ""
   End If
   On Error GoTo DiaErr1
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("engrt02")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Includes='Includes " & cmbRte & "...'"
'   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDesc"
   aFormulaName.Add "ShowComments"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Includes " & CStr(cmbRte) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.value
   aFormulaValue.Add optCmt.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RthdTable.RTREF} Like '" & sRout & "*' "
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

'   MDISect.Crw.SelectionFormula = "{RthdTable.RTREF} Like '" & sRout & "*' "
'   If optDsc.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPHDR.1.0;T;;;"
'   End If
'   If OptCmt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(1) = "GROUPFTR.1.0;F;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPFTR.1.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(1) = "GROUPFTR.1.0;T;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPFTR.1.1;T;;;"
'   End If
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Save by Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "rt02", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.value = Val(Left(sOptions, 1))
      optCmt.value = Val(Mid(sOptions, 2, 1))
   End If
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = RTrim(optDsc.value) _
              & RTrim(optCmt.value)
   SaveSetting "Esi2000", "EsiEngr", "rt02", Trim(sOptions)
   
End Sub
