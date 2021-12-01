VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form diaSHp07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Shift Codes"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "diaSHp07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaSHp07.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3120
      Top             =   240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1530
      FormDesignWidth =   5790
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Shift Code"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "diaSHp07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOnLoad As Byte


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT SFCODE FROM SfcdTable "
   LoadComboBox cmbCde, -1
   If cmbCde.ListCount > 0 Then
      cmbCde = cmbCde.List(0)
      
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Sub cmbCde_LostFocus()
    If Len(cmbCde) = 0 Then cmbCde = "ALL"
End Sub

Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
   Show
End Sub


Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaSHp07 = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub optDis_Click()
    If Len(cmbCde) = 0 Then cmbCde = "ALL"
    PrintReport
End Sub

Private Sub optPrn_Click()
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   PrintReport
   
End Sub


Private Sub PrintReport()
   Dim sCode As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   If cmbCde <> "ALL" Then sCode = Left(cmbCde, 2) Else sCode = ""
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaValue.Add CStr("'Requested By: " & sFacility & "'")
   aFormulaValue.Add CStr("'Includes " & cmbCde & "...'")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admsh01")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{SfcdTable.SFCODE} Like '" & sCode & "*' "
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


