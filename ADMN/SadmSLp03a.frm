VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SadmSLp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Persons Report"
   ClientHeight    =   2640
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
   ScaleHeight     =   2640
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SadmSLp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbSlp 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Salesperson (4 Char)"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   4
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
         Picture         =   "SadmSLp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SadmSLp03a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
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
      FormDesignHeight=   2640
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Detail"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person(s)"
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   6
      Top             =   1080
      Width           =   1665
   End
End
Attribute VB_Name = "SadmSLp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/9/05 corrected FillSales query
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbSlp_LostFocus()
   cmbSlp = CheckLen(cmbSlp, 4)
   If Len(cmbSlp) = 0 Then cmbSlp = "ALL"
   
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
   If bOnLoad Then FillSales
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub FillSales()
   'Dim RdoCmb As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT SPNUMBER FROM SprsTable "
   LoadComboBox cmbSlp, -1
   cmbSlp = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillsales"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
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
   Set SadmSLp03a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim iList As Integer
   Dim sPerson As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   If cmbSlp <> "ALL" Then sPerson = cmbSlp Else sPerson = ""
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDet"
   aFormulaValue.Add CStr("'" & sFacility & "'")
   aFormulaValue.Add CStr("'" & cmbSlp & "...'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   aFormulaValue.Add optDet
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admco03")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{SprsTable.SPNUMBER} Like '" & sPerson & "*' "
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
   Dim iList As Integer
   Dim sPerson As String
   MouseCursor 13
   'SetMdiReportsize MDISect
   
   If cmbSlp <> "ALL" Then sPerson = cmbSlp Else sPerson = ""
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "Includes='Includes " & cmbSlp & "... '"
   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sCustomReport = GetCustomReport("admco03")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   If optDet Then
      For iList = 0 To 11
         MDISect.Crw.SectionFormat(iList) = "DETAIL.0." & Trim(str(iList)) & ";T;;;;"
      Next
   Else
      For iList = 0 To 11
         MDISect.Crw.SectionFormat(iList) = "DETAIL.0." & Trim(str(iList)) & ";F;;;;"
      Next
   End If
   MDISect.Crw.SelectionFormula = "{SprsTable.SPNUMBER} Like '" & sPerson & "*' "
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
   Dim sOptions As String
   sOptions = Trim(str(optDet.value)) & cmbSlp
   SaveSetting "Esi2000", "EsiAdmn", "co03", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiAdmn", "co03", sOptions)
   optDet.value = Val(Left(sOptions, 1))
   cmbSlp = Mid(sOptions, 2, 4)
   If cmbSlp = "" Then cmbSlp = "ALL"
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If Len(cmbSlp) = 0 Then cmbSlp = "ALL"
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   If Len(cmbSlp) = 0 Then cmbSlp = "ALL"
   PrintReport
   
End Sub
