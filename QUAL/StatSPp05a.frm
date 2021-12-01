VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processes By Part Number"
   ClientHeight    =   3180
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDat 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optPrc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   2200
      Width           =   735
   End
   Begin VB.CheckBox optRes 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Combo Box Is Filled With Parts Containing Key Dimensions"
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "StatSPp05a.frx":07AE
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
         Picture         =   "StatSPp05a.frx":092C
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
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3180
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Sources"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Process ID's"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   2200
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reasoning Codes"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "StatSPp05a"
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

Private Sub cmbPrt_Click()
   GetSpcPart
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      GetSpcPart
   Else
      cmbPrt = "ALL"
      lblDsc = "All Parts Selected.."
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
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,KEYREF FROM " _
          & "PartTable,RjkyTable WHERE PARTREF=KEYREF ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
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
   Set StatSPp05a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sPart As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
   sCustomReport = GetCustomReport("quasp05")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowRes"
   aFormulaName.Add "ShowPrc"
   aFormulaName.Add "ShowDat"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbPrt) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optRes.value
   aFormulaValue.Add optPrc.value
   aFormulaValue.Add optDat.value
   sSql = "{PartTable.PARTREF} Like '" & sPart & "*' "
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
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
   Dim sPart As String
   MouseCursor 13
   On Error GoTo DiaErr1
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("quasp05")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='" & cmbPrt & "...'"
   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   sSql = "{PartTable.PARTREF} Like '" & sPart & "*' "
   If optRes.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
   End If
   If optPrc.value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.2.0;F;;;"
      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.2.1;F;;;"
   Else
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.2.0;T;;;"
      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.2.1;T;;;"
   End If
   If optDat.value = vbChecked Then
      MdiSect.Crw.SectionFormat(4) = "GROUPFTR.1.0;T;;;"
      MdiSect.Crw.SectionFormat(5) = "GROUPFTR.1.1;T;;;"
      MdiSect.Crw.SectionFormat(6) = "GROUPFTR.1.2;T;;;"
      MdiSect.Crw.SectionFormat(7) = "GROUPFTR.1.3;T;;;"
      MdiSect.Crw.SectionFormat(8) = "GROUPFTR.1.4;T;;;"
      MdiSect.Crw.SectionFormat(9) = "GROUPFTR.1.5;T;;;"
      MdiSect.Crw.SectionFormat(10) = "GROUPFTR.1.6;T;;;"
      MdiSect.Crw.SectionFormat(11) = "GROUPFTR.1.7;T;;;"
      MdiSect.Crw.SectionFormat(12) = "GROUPFTR.1.8;T;;;"
      MdiSect.Crw.SectionFormat(13) = "GROUPFTR.1.9;T;;;"
      MdiSect.Crw.SectionFormat(14) = "GROUPFTR.1.10;T;;;"
      MdiSect.Crw.SectionFormat(15) = "GROUPFTR.1.11;T;;;"
      MdiSect.Crw.SectionFormat(16) = "GROUPFTR.1.12;T;;;"
      MdiSect.Crw.SectionFormat(17) = "GROUPFTR.1.13;T;;;"
      MdiSect.Crw.SectionFormat(18) = "GROUPFTR.1.14;T;;;"
      MdiSect.Crw.SectionFormat(19) = "GROUPFTR.1.15;T;;;"
      MdiSect.Crw.SectionFormat(20) = "GROUPFTR.1.16;T;;;"
      MdiSect.Crw.SectionFormat(21) = "GROUPFTR.1.17;T;;;"
      MdiSect.Crw.SectionFormat(22) = "GROUPFTR.1.18;T;;;"
      MdiSect.Crw.SectionFormat(23) = "GROUPFTR.1.19;T;;;"
   Else
      MdiSect.Crw.SectionFormat(4) = "GROUPFTR.1.0;F;;;"
      MdiSect.Crw.SectionFormat(5) = "GROUPFTR.1.1;F;;;"
      MdiSect.Crw.SectionFormat(6) = "GROUPFTR.1.2;F;;;"
      MdiSect.Crw.SectionFormat(7) = "GROUPFTR.1.3;F;;;"
      MdiSect.Crw.SectionFormat(8) = "GROUPFTR.1.4;F;;;"
      MdiSect.Crw.SectionFormat(9) = "GROUPFTR.1.5;F;;;"
      MdiSect.Crw.SectionFormat(10) = "GROUPFTR.1.6;F;;;"
      MdiSect.Crw.SectionFormat(11) = "GROUPFTR.1.7;F;;;"
      MdiSect.Crw.SectionFormat(12) = "GROUPFTR.1.8;F;;;"
      MdiSect.Crw.SectionFormat(13) = "GROUPFTR.1.9;F;;;"
      MdiSect.Crw.SectionFormat(14) = "GROUPFTR.1.10;F;;;"
      MdiSect.Crw.SectionFormat(15) = "GROUPFTR.1.11;F;;;"
      MdiSect.Crw.SectionFormat(16) = "GROUPFTR.1.12;F;;;"
      MdiSect.Crw.SectionFormat(17) = "GROUPFTR.1.13;F;;;"
      MdiSect.Crw.SectionFormat(18) = "GROUPFTR.1.14;F;;;"
      MdiSect.Crw.SectionFormat(19) = "GROUPFTR.1.15;F;;;"
      MdiSect.Crw.SectionFormat(20) = "GROUPFTR.1.16;F;;;"
      MdiSect.Crw.SectionFormat(21) = "GROUPFTR.1.17;F;;;"
      MdiSect.Crw.SectionFormat(22) = "GROUPFTR.1.18;F;;;"
      MdiSect.Crw.SectionFormat(23) = "GROUPFTR.1.19;F;;;"
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
   sOptions = Trim(str(optRes.value)) _
              & Trim(str(optPrc.value)) & Trim(str(optDat))
   SaveSetting "Esi2000", "EsiQual", "sp05", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "sp05", Trim(sOptions))
   If Len(Trim(sOptions)) Then
      optRes.value = Val(Left(sOptions, 1))
      optPrc.value = Val(Mid(sOptions, 2, 1))
      optDat.value = Val(Right(sOptions, 1))
   Else
      optRes.value = vbChecked
      optPrc.value = vbChecked
      optDat.value = vbChecked
   End If
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 7) = "*** Not" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optDat_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub GetSpcPart()
   Dim RdoPrt As ADODB.Recordset
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
          & "WHERE PARTREF='" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         lblDsc = "" & Trim(!PADESC)
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "Series Of Parts Selected."
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optRes_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
