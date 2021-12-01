VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSp17a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manifest Printing"
   ClientHeight    =   2565
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2565
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp17a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEndDte 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbMan 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Pack Slips"
      Top             =   1680
      Width           =   1555
   End
   Begin VB.ComboBox txtStartDte 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp17a.frx":07AE
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
         Picture         =   "PackPSp17a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2565
      FormDesignWidth =   7005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   11
      Top             =   1080
      Width           =   1290
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS End Date"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Manifest"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS Start Date"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   650
      Width           =   1065
   End
End
Attribute VB_Name = "PackPSp17a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/3/05 New
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
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
   
      If (txtStartDte.Text = "") Then txtStartDte = Format(ES_SYSDATE, "mm/dd/yy")
      If (txtEndDte.Text = "") Then txtEndDte = Format(ES_SYSDATE, "mm/dd/yy")
   
      FillManifestNum
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub PrintReport()
   Dim sBeg As String
   Dim sEnd As String
   Dim sCust As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim strASN As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   strASN = cmbMan.Text
   If (strASN = "") Then
      MsgBox ("Please select Manifest Number")
      Exit Sub
   End If
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slesh05a")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   sSql = "{EsReportASNManifest.PSSHIPNO} = " & Val(strASN)
   
   
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


Private Sub FillManifestNum()
         
   Dim strStartDate As String
   Dim strEndDate As String
   
   strStartDate = txtStartDte.Text
   strEndDate = txtEndDte.Text
   If (strStartDate <> "ALL" And strEndDate <> "ALL") Then
      sSql = "SELECT DISTINCT PSSHIPNO From PshdTable " _
            & " WHERE PshdTable.PSDATE BETWEEN '" & strStartDate _
            & "' AND '" & strEndDate & "' AND PSSHIPNO <> 0"
   Else
      sSql = "SELECT DISTINCT PSSHIPNO From PshdTable"
   End If
   
   LoadComboBox cmbMan, -1
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
End Sub

Private Sub txtEndDte_LostFocus()
   If txtEndDte = "" Then txtEndDte = "ALL"
   If txtEndDte <> "ALL" Then txtEndDte = CheckDate(txtEndDte)
   
   FillManifestNum
End Sub

Private Sub txtStartDte_LostFocus()
   If txtStartDte = "" Then txtStartDte = "ALL"
   If txtStartDte <> "ALL" Then txtStartDte = CheckDate(txtStartDte)
   
   FillManifestNum
End Sub

Private Sub SaveOptions()
End Sub

Private Sub txtStartDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub GetOptions()
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set PackPSp17a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub


Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
