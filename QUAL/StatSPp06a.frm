VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Numbers By Family ID"
   ClientHeight    =   3060
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
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optRes 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   735
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
         Picture         =   "StatSPp06a.frx":07AE
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
         Picture         =   "StatSPp06a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbFam 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   1080
      Width           =   1815
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Family ID"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "StatSPp06a"
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

Private Sub GetFamily()
   Dim RdoFam As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FAMREF,FAMID,FAMDESC,FAMNOTES FROM " _
          & "RjfmTable WHERE FAMREF='" & Compress(cmbFam) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFam, ES_FORWARD)
   If bSqlRows Then
      With RdoFam
         cmbFam = "" & Trim(!FAMID)
         lblDsc = "" & Trim(!FAMDESC)
         ClearResultSet RdoFam
      End With
   Else
      lblDsc = "Range Of Families Selected"
   End If
   Set RdoFam = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getfamily"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbFam_Click()
   GetFamily
   
End Sub

Private Sub cmbFam_LostFocus()
   cmbFam = CheckLen(cmbFam, 15)
   If Len(cmbFam) Then
      GetFamily
   Else
      cmbFam = "ALL"
      lblDsc = "All Families Selected"
   End If
   
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
   sSql = "Qry_FillSPFamily"
   LoadComboBox cmbFam
   If cmbFam.ListCount > 0 Then cmbFam = cmbFam.List(0)
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
   Set StatSPp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sFam As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   On Error GoTo DiaErr1
   If cmbFam <> "ALL" Then sFam = Compress(cmbFam)
   sCustomReport = GetCustomReport("quasp06")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowRes"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbFam) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optRes.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   sSql = "{RjfmTable.FAMREF} LIKE '" & sFam & "*' "
   cCRViewer.SetReportSelectionFormula (sSql)
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
   Dim sFam As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   If cmbFam <> "ALL" Then sFam = Compress(cmbFam)
   sCustomReport = GetCustomReport("quasp06")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='" & cmbFam & "...'"
   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sSql = "{RjfmTable.FAMREF} LIKE '" & sFam & "*' "
   If optRes.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
      MdiSect.Crw.SectionFormat(1) = "GROUPFTR.1.0;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
      MdiSect.Crw.SectionFormat(1) = "GROUPFTR.1.0;T;;;"
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
   sOptions = Trim(str(optRes.value))
   SaveSetting "Esi2000", "EsiQual", "sp06", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "Esiqual", "sp06", Trim(sOptions))
   If Len(sOptions) > 0 Then
      optRes.value = Val(sOptions)
   Else
      optRes.value = vbChecked
   End If
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub optRes_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
