VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Of Team Members"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtDpt 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1560
      Width           =   1605
   End
   Begin VB.ComboBox cmbMem 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Leading Chars"
      Top             =   840
      Width           =   1875
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Or Leading Chars"
      Top             =   1200
      Width           =   860
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "StatSPp01a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "StatSPp01a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   6885
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Comments"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department(s)"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Team Member(s)"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division(s)"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "StatSPp01a"
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


Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If Trim(cmbDiv) = "" Then cmbDiv = "ALL"
   
End Sub


Private Sub cmbMem_LostFocus()
   cmbMem = CheckLen(cmbMem, 15)
   If Trim(cmbMem) = "" Then cmbMem = "ALL"
   
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


Private Sub FillCombo()
   On Error GoTo DiaErr1
   If cmbDiv = "" Then cmbDiv = "ALL"
   AddComboStr cmbMem.hwnd, "ALL"
   sSql = "Qry_FillSPTeam"
   LoadComboBox cmbMem, -1
   If cmbMem.ListCount > 0 Then cmbMem = cmbMem.List(0)
   If Trim(txtDpt) = "" Then txtDpt = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      AddComboStr cmbDiv.hwnd, "ALL"
      FillDivisions
      FillCombo
      bOnLoad = 0
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
   Set StatSPp01a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sMem As String
   Dim sDiv As String
   Dim sDpt As String
   MouseCursor 13
   On Error GoTo DiaErr1
   
   
   If Trim(cmbMem) <> "ALL" Then sMem = Compress(cmbMem)
   If Trim(cmbDiv) <> "ALL" Then sDiv = Trim(cmbDiv)
   If Trim(txtDpt) <> "ALL" Then sDpt = Trim(txtDpt)
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("quasp01")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowComments"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Member(s) " & Trim(cmbMem) _
                        & ", Division(s) " & Trim(cmbDiv) & ", Deptartment()s " _
                        & Trim(txtDpt) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add OptCmt.value
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RjtmTable.TMMID} LIKE '" & sMem & "*' " _
          & "AND {RjtmTable.TMMDIVISION} LIKE '" & sDiv & "*' " _
          & "AND {RjtmTable.TMMDEPARTMENT} LIKE '" & sDpt & "*'"
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
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
   Dim sMem As String
   Dim sDiv As String
   Dim sDpt As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If Trim(cmbMem) <> "ALL" Then sMem = Compress(cmbMem)
   If Trim(cmbDiv) <> "ALL" Then sDiv = Trim(cmbDiv)
   If Trim(txtDpt) <> "ALL" Then sDpt = Trim(txtDpt)
   
   
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("quasp01")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='Member(s) " & Trim(cmbMem) _
                        & ", Division(s) " & Trim(cmbDiv) & ", Deptartment()s " _
                        & Trim(txtDpt) & "...'"
   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sSql = "{RjtmTable.TMMID} LIKE '" & sMem & "*' " _
          & "AND {RjtmTable.TMMDIVISION} LIKE '" & sDiv & "*' " _
          & "AND {RjtmTable.TMMDEPARTMENT} LIKE '" & sDpt & "*'"
   If OptCmt.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
   End If
   MdiSect.Crw.SelectionFormula = sSql
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
   Dim sIdn As String * 15
   Dim sDiv As String * 3
   Dim sDpt As String * 15
   sIdn = cmbMem
   sDiv = cmbDiv
   sDpt = txtDpt
   sOptions = sIdn & sDiv & sDpt & Trim(str(OptCmt.value))
   SaveSetting "Esi2000", "EsiQual", "sp01", Trim(sOptions)
   
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "sp01", Trim(sOptions))
   If Len(Trim(sOptions)) Then
      cmbMem = Trim(Mid(sOptions, 1, 15))
      cmbDiv = Trim(Mid(sOptions, 16, 3))
      txtDpt = Trim(Mid(sOptions, 19, 15))
      OptCmt.value = Val(Mid(sOptions, 34, 1))
   Else
      OptCmt.value = vbChecked
   End If
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtDpt_LostFocus()
   txtDpt = CheckLen(txtDpt, 15)
   If Trim(txtDpt) = "" Then txtDpt = "ALL"
   
End Sub
