VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BompBMp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts List Report"
   ClientHeight    =   3510
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkShowBomComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   3060
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "BompBMp01a.frx":07AE
      Height          =   350
      Left            =   5400
      Picture         =   "BompBMp01a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Parts List for Part and Revision"
      Top             =   1080
      Width           =   350
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BompBMp01a.frx":15FA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BompBMp01a.frx":1784
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision (Blank For Default)"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   2460
      Width           =   735
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.ComboBox cmbPls 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Only Parts With A Parts List"
      Top             =   1110
      Width           =   3345
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
      FormDesignHeight=   3510
      FormDesignWidth =   7515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Comments"
      Height          =   165
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   3060
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   288
      Index           =   5
      Left            =   5880
      TabIndex        =   16
      Top             =   1110
      Width           =   612
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6840
      TabIndex        =   15
      Top             =   1110
      Width           =   300
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   14
      Top             =   1440
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   168
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts List Comments"
      Height          =   165
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Part Desc"
      Height          =   165
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2460
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts List"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1110
      Width           =   1815
   End
End
Attribute VB_Name = "BompBMp01a"
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
Dim bGoodList As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "pl01", sOptions) & "00000000"
   optDsc.Value = Mid(sOptions, 1, 1)
   OptCmt.Value = Mid(sOptions, 2, 1)
   chkShowBomComments = Mid(sOptions, 3, 1)
End Sub


Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = optDsc & OptCmt & Me.chkShowBomComments
   SaveSetting "Esi2000", "EsiEngr", "pl01", Trim(sOptions)
   
End Sub

Private Sub GetPartRevision()
   Dim RdoRev As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF='" & Compress(cmbPls) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev)
   If bSqlRows Then
      With RdoRev
         cmbRev = "" & Trim(!PABOMREV)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = "" & Trim(str(!PALEVEL))
         ClearResultSet RdoRev
      End With
   Else
      cmbRev = ""
   End If
   Set RdoRev = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartrev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub





Private Sub cmbPls_Click()
   FillBomhRev cmbPls
   GetPartRevision
   
End Sub

Private Sub cmbPls_KeyUp(KeyCode As Integer, Shift As Integer)
   cmbPls = CheckLen(cmbPls, 30)
   FillBomhRev cmbPls
   GetPartRevision
   
End Sub

Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   FillBomhRev cmbPls
   GetPartRevision
   
End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   
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



Private Sub cmdVew_Click()
   If cmdVew Then
      ViewBomTree.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   MouseCursor 0
   If bOnLoad = 1 Then
      FillCombo
      'If cUR.CurrentPart <> "" Then cmbPls = cUR.CurrentPart
      FillBomhRev cmbPls
      If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
      bOnLoad = 0
   End If
   MDISect.lblBotPanel = Caption
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
   GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveCurrentSelections
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   SaveCurrentSelections
   FormUnload
   Set BompBMp01a = Nothing
   
End Sub
Private Sub PrintReport()
   MouseCursor 13
   Dim sRev As String
   Dim sPls As String
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   cmbRev = Compress(cmbRev)
   sPls = Compress(cmbPls)
   sRev = cmbRev
   'SetMdiReportsize MDISect
   On Error GoTo Eng01
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowBomComments"
   aFormulaName.Add "ShowComments"
   aFormulaName.Add "ShowDescription"
   
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(Me.chkShowBomComments) & "'")
   aFormulaValue.Add OptCmt
   aFormulaValue.Add optDsc

   'MDISect.Crw.Formulas(0) = "RequestBy='Requested By: " & sInitials & "'"
   'MDISect.Crw.Formulas(1) = "ShowBomComments='" & Me.chkShowBomComments & "'"
   sCustomReport = GetCustomReport("engbm01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{PartTable.PARTREF}='" & sPls & "' " _
                                  & "AND {BmhdTable.BMHREV}='" & sRev & "' "
                                  
                                  
   'MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   'MDISect.Crw.SelectionFormula = "{PartTable.PARTREF}='" & sPls & "' " _
                                  & "AND {BmhdTable.BMHREV}='" & sRev & "' "
'   If OptCmt.value = 0 Then
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPFTR.0.1;F;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPFTR.0.2;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPFTR.0.1;T;;;"
'      MDISect.Crw.SectionFormat(2) = "GROUPFTR.0.2;T;;;"
'   End If
'   If optDsc.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(3) = "DETAIL.0.0;F;;;"
'      MDISect.Crw.SectionFormat(4) = "DETAIL.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(3) = "DETAIL.0.0.;T;;;"
'      MDISect.Crw.SectionFormat(4) = "DETAIL.0.1.;T;;;"
'   End If
'   SetCrystalAction Me

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
   
Eng01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   Dim bGoodRout As Byte
   MouseCursor 13
   bGoodList = GetList()
   If Not bGoodList Then
      MouseCursor 0
      MsgBox "Parts List Wasn't Found.", vbExclamation, Caption
      Exit Sub
   Else
      PrintReport
   End If
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Function GetList() As Byte
   Dim RdoPls As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BMHREF,BMHREV FROM " _
          & "BmhdTable WHERE BMHREF='" & Compress(cmbPls) & "' " _
          & "AND BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls)
   If bSqlRows Then
      cmbRev = "" & Trim(RdoPls!BMHREV)
      GetList = True
   Else
      GetList = False
   End If
   Set RdoPls = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optPrn_Click()
   Dim bGoodRout As Byte
   MouseCursor 13
   bGoodList = GetList()
   If Not bGoodList Then
      MouseCursor 0
      MsgBox "Parts List Wasn't Found.", vbExclamation, Caption
      Exit Sub
   Else
      PrintReport
   End If
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BMASSYPART,PARTREF,PARTNUM FROM " _
          & "BmplTable,PartTable WHERE BMASSYPART=PARTREF " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbPls, 1
   If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
