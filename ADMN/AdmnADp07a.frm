VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnADp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Report Settings"
   ClientHeight    =   2715
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbSection 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Class From List"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "AdmnADp07a.frx":07AE
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
         Picture         =   "AdmnADp07a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2715
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label S 
      BackStyle       =   0  'Transparent
      Caption         =   "Section (s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "AdmnADp07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillSections()
   Dim RdoCmb As ADODB.Recordset
   Dim sLastSection As String
   On Error GoTo DiaErr1
   'Distinct doesn't work here
   sSql = "SELECT REPORT_SECTION FROM CustomReports " _
          & "ORDER BY REPORT_INDEX"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         AddComboStr cmbSection.hwnd, "ALL"
         Do Until .EOF
            If sLastSection <> Trim(!REPORT_SECTION) Then
               AddComboStr cmbSection.hwnd, "" & Trim(!REPORT_SECTION)
               sLastSection = "" & Trim(!REPORT_SECTION)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbSection.ListCount > 0 Then
      cmbSection = cmbSection.List(0)
   Else
      MsgBox "Custom Reports Has Not Been Installed.", _
         vbInformation, Caption
      Unload Me
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillsections"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbSection_LostFocus()
   If Trim(cmbSection) = "" Then cmbSection = "ALL"
   
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
   If bOnLoad = 1 Then FillSections
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
   Set AdmnADp07a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sClass As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbSection <> "ALL" Then sClass = cmbSection
'   SetMdiReportsize MDISect
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admco07")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Includes='" & cmbSection & "...'"
'   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'" & sFacility & "'")
   aFormulaValue.Add CStr("'" & cmbSection & "...'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{CustomReports.REPORT_SECTION} LIKE '" & sClass & "*' "
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
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
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sClass As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If cmbSection <> "ALL" Then sClass = cmbSection
   'SetMdiReportsize MDISect
   sCustomReport = GetCustomReport("admco07")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "Includes='" & cmbSection & "...'"
   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sSql = "{CustomReports.REPORT_SECTION} LIKE '" & sClass & "*' "
   MDISect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
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
   
End Sub

Private Sub GetOptions()
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub
