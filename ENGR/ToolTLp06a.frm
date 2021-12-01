VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ToolTLp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   2265
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
   ScaleHeight     =   2265
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbTol 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select - Leading Chars"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.ComboBox cmbCls 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "2"
      ToolTipText     =   "Select - Leading Chars"
      Top             =   1080
      Width           =   1935
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
         Picture         =   "ToolTLp06a.frx":07AE
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
         Picture         =   "ToolTLp06a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2265
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   0
      Left            =   5640
      TabIndex        =   9
      Top             =   1440
      Width           =   1428
   End
   Begin VB.Label T 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Number(s)"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5640
      TabIndex        =   7
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label T 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Class(es)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "ToolTLp06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'3/8/05 Fixed query
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT TOOL_CLASS FROM TlnhdTableNew"
   LoadComboBox cmbCls, -1
   If cmbCls.ListCount > 0 Then cmbCls = cmbCls.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCls_Click()
   FillTools
   
End Sub


Private Sub cmbCls_LostFocus()
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   FillTools
   
End Sub

Private Sub cmbTol_LostFocus()
   If Trim(cmbTol) = "" Then cmbTol = "ALL"
   
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
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
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
   FormUnload
   Set ToolTLp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sClass As String
   Dim sPart As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   If cmbCls <> "ALL" Then sClass = UCase$(cmbCls)
   If cmbTol <> "ALL" Then sPart = Compress(cmbTol)
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Includes Class(es) " & CStr(cmbCls _
                        & ", Tool(s) " & cmbTol) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("engtl06")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   sSql = "UpperCase({TlnhdTableNew.TOOL_CLASS}) LIKE '" & sClass & "*' " _
          & "AND {TlnhdTableNew.TOOL_NUM} LIKE '" & sPart & "*' "
   
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
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   
End Sub

Private Sub GetOptions()
   On Error Resume Next
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub FillTools()
   Dim sClass As String
   On Error GoTo DiaErr1
   cmbTol.Clear
   If cmbCls <> "ALL" Then sClass = cmbCls
   sSql = "SELECT TOOL_NUM FROM TlnhdTableNew WHERE TOOL_CLASS LIKE '" _
          & sClass & "%' ORDER BY TOOL_NUM"
   LoadComboBox cmbTol, -1
   If cmbTol.ListCount > 0 Then cmbTol = cmbTol.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "filltools"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
