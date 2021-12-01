VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packing Slip Edit"
   ClientHeight    =   3615
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7170
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbEps 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Enter Pack Slip Or Select From List"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cmbBps 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Enter Pack Slip Or Select From List"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PackPSp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp03a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.CheckBox optRem 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox optFet 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3615
      FormDesignWidth =   7170
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Feature Options"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   2640
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(If Different)"
      Height          =   195
      Index           =   5
      Left            =   3480
      TabIndex        =   13
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending PS Number"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1380
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting PS Number"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1000
      Width           =   1905
   End
End
Attribute VB_Name = "PackPSp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/20/05 Corrected missing Qry_FillPackSlipsNotPrinted
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetLastPackslip()
'   Dim RdoGet As ADODB.Recordset
'   On Error GoTo DiaErr1
'   sSql = "SELECT CURPSNUMBER FROM ComnTable WHERE COREF=1"
'   bSqlRows = GetDataSet(RdoGet, ES_FORWARD)
'   If bSqlRows Then
'      On Error Resume Next
'      With RdoGet
'         If Val(!CURPSNUMBER) > 0 Then
'            cmbBps = "PS" & Trim(!CURPSNUMBER)
'         Else
'            cmbBps = "" & Trim(!CURPSNUMBER)
'         End If
'         cmbEps = cmbBps
'         ClearResultSet RdoGet
'      End With
'   End If
'   Set RdoGet = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getpacksl"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me

   Dim ps As New ClassPackSlip
   cmbBps = ps.GetLastPackSlipNumber
End Sub

Private Sub FillCombo()
'   Dim RdoCmb As ADODB.Recordset
'   Dim iList As Integer
'   cmbBps.Clear
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   sSql = "Qry_FillPackSlipsNotPrinted '" & DateAdd("yyyy", -2, Now) & "'"
'   bSqlRows = GetDataSet(RdoCmb)
'   If bSqlRows Then
'      With RdoCmb
'         Do Until .EOF
'            AddComboStr cmbBps.hWnd, "" & Trim(!PsNumber)
'            AddComboStr cmbEps.hWnd, "" & Trim(!PsNumber)
'            .MoveNext
'         Loop
'         ClearResultSet RdoCmb
'      End With
'   Else
'      MouseCursor 0
'      MsgBox "No Sales Orders Where Found.", vbInformation, Caption
'      Exit Sub
'   End If
'   Set RdoCmb = Nothing
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
   
   Dim ps As New ClassPackSlip
   ps.FillPSComboUnprinted cmbBps
   ps.FillPSComboUnprinted cmbEps
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optDsc.Value) _
              & RTrim(optExt.Value) _
              & RTrim(optCmt.Value) _
              & RTrim(optRem.Value) _
              & RTrim(optFet.Value)
   SaveSetting "Esi2000", "EsiSale", "sh03", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "Sh03", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.Value = Val(Left(sOptions, 1))
      optExt.Value = Val(Mid(sOptions, 2, 1))
      optCmt.Value = Val(Mid(sOptions, 3, 1))
      optRem.Value = Val(Mid(sOptions, 4, 1))
      optFet.Value = Val(Mid(sOptions, 5, 1))
   End If
   
End Sub

Private Sub cmbBps_LostFocus()
   cmbBps = CheckLen(cmbBps, 8)
   If Len(cmbBps) = 0 Then cmbBps = "ALL"
   
End Sub


Private Sub cmbEps_LostFocus()
   cmbEps = CheckLen(cmbEps, 8)
   If Len(cmbEps) = 0 Then cmbEps = "ALL"
   
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
   If bOnLoad Then
      FillCombo
      GetLastPackslip
      bOnLoad = 0
   End If
   MdiSect.lblBotPanel = Caption
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
   Set PackPSp03a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim sBegPs As String
    Dim sEndPs As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   If Len(cmbBps) = 0 Or cmbBps = "ALL" Then
      sBegPs = "0000000"
      cmbBps = "ALL"
   Else
      sBegPs = cmbBps
   End If
   If Len(cmbEps) = 0 Or cmbEps = "ALL" Then
      sEndPs = "zzzzzzz"
      cmbEps = "ALL"
   Else
      sEndPs = cmbEps
   End If
   MouseCursor 13
   On Error GoTo DiaErr1
   aFormulaName.Add "ShowRemarks"
   aFormulaName.Add "ShowDescription"
   aFormulaName.Add "ShowExDescription"
   aFormulaName.Add "ShowComments"
   aFormulaValue.Add optRem.Value
   aFormulaValue.Add optDsc.Value
   aFormulaValue.Add optExt.Value
   aFormulaValue.Add optCmt.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleps03")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   If cmbBps = "ALL" Then cmbBps = ""
   If cmbEps = "ALL" Then cmbEps = ""
   sSql = "{PshdTable.PSNUMBER} in '" & sBegPs & "' to '" _
          & "" & sEndPs & "'"
   
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

Private Sub optCmt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   MouseCursor 11
   PrintReport
   
End Sub

Private Sub optDsc_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFet_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optFet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   MouseCursor 11
   PrintReport
   
End Sub

Private Sub optRem_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optRem_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
