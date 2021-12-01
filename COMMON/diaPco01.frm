VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPco01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Divisions (Report)"
   ClientHeight    =   2460
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   1080
      Width           =   860
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
         Picture         =   "diaPco01.frx":0000
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
         Picture         =   "diaPco01.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   2
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
      PictureUp       =   "diaPco01.frx":0308
      PictureDn       =   "diaPco01.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2460
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Width           =   3825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division(s) "
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "diaPco01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If Len(cmbDiv) = 0 Then cmbDiv = "ALL"
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs907"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillDivisions Me
      bOnLoad = 0
      cmbDiv = "ALL"
   End If
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
   Set diaPco01 = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sDivision As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If cmbDiv <> "ALL" Then sDivision = Trim(cmbDiv)
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Includes" & CStr(cmbDiv) & "...'")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("admco01")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

    sSql = "{CdivTable.DIVREF} LIKE '" & sDivision & "*' "
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
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

Private Sub PrintReport1()
   Dim sDivision As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If cmbDiv <> "ALL" Then sDivision = Trim(cmbDiv)
   'SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='Includes " & cmbDiv & "...'"
   sCustomReport = GetCustomReport("admco01")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{CdivTable.DIVREF} LIKE '" & sDivision & "*' "
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiAdmn", "co01", Trim(cmbDiv)
   
End Sub

Private Sub GetOptions()
   cmbDiv = GetSetting("Esi2000", "EsiAdmn", "co01", cmbDiv)
   If cmbDiv = "" Then cmbDiv = "ALL"
   
End Sub

Private Sub optDis_Click()
   If Trim(cmbDiv) = "" Then cmbDiv = "ALL"
   If cmbDiv <> "ALL" Then cmbDiv = Left(cmbDiv, 4)
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   If Trim(cmbDiv) = "" Then cmbDiv = "ALL"
   If cmbDiv <> "ALL" Then cmbDiv = Left(cmbDiv, 4)
   PrintReport
   
End Sub
