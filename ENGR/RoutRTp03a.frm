VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routings By Part Number"
   ClientHeight    =   2730
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2730
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Routings"
      Top             =   1200
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "RoutRTp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "RoutRTp03a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2730
      FormDesignWidth =   7095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "RoutRTp03a"
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
   If bOnLoad = 1 Then FillCombo
   bOnLoad = 0
   MDISect.lblBotPanel = Caption
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
   cmbRte = "ALL"
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set RoutRTp03a = Nothing
   
End Sub

Private Sub PrintReport()
   MouseCursor 13
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
'   SetMdiReportsize MDISect
   
   Dim sRout As String
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
   sCustomReport = GetCustomReport("engrt03")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Includes='Includes " & cmbRte & "...'"
'   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Includes " & CStr(cmbRte) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{PartTable.PARTREF} Like '" & sRout & "*' "
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
'   MDISect.Crw.SelectionFormula = "{PartTable.PARTREF} Like '" & sRout & "*' "
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PAROUTING,RTREF,RTNUM " _
          & "FROM PartTable,RthdTable WHERE PARTREF=RTREF " _
          & "ORDER BY RTREF"
   LoadComboBox cmbRte, 3
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub
