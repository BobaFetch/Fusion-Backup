VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form jevShopSHp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "jevSHp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "jevSHp01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optFrom 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "jevSHp01a.frx":0938
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
         Picture         =   "jevSHp01a.frx":0AC2
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
      Left            =   6240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Part Number"
      Top             =   1320
      Width           =   3545
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6960
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2850
      FormDesignWidth =   7515
   End
   Begin VB.Label lblName 
      Caption         =   "Jevco Custom Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblQty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSta 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   11
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblTyp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Top             =   1680
      Width           =   300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type/Status"
      Height          =   252
      Index           =   15
      Left            =   5040
      TabIndex        =   9
      Top             =   1680
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "jevShopSHp01a"
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
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

'Dim DbDoc   As Recordset 'Jet
'Dim DbPls   As Recordset 'Jet

Dim bGoodDocs As Boolean
Dim bGoodPlst As Boolean

Dim bCanceled As Byte
Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Byte

Dim sRunPkstart As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   lblPrinter = GetSetting("Esi2000", "EsiProd", "sh01Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiProd", "sh01Printer", lblPrinter
   
End Sub




Private Sub cmbPrt_Click()
   bGoodPart = GetRuns()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled = 0 Then bGoodPart = GetRuns()
   
End Sub


Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo Psh01
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   If Dir(sReportPath & "jevprdsh01.rpt") <> "" Then
      sCustomReport = GetCustomReport("jevprdsh01")
   Else
      sCustomReport = GetCustomReport("prdsh01")
   End If
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   
   sPartNumber = Compress(cmbPrt)
   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(cmbRun) & " "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   
   MouseCursor 0
   Exit Sub
   
Psh01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
   
End Sub


'Private Sub PrintReport1()
'   MouseCursor 13
'   On Error GoTo Psh01
'   SetMdiReportsize MDISect
'   If Dir(sReportPath & "jevprdsh01.rpt") <> "" Then
'      sCustomReport = GetCustomReport("jevprdsh01")
'      MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   Else
'      sCustomReport = GetCustomReport("prdsh01")
'      MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   End If
'   sPartNumber = Compress(cmbPrt)
'   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
'          & "AND {RunsTable.RUNNO}=" & Trim(cmbRun) & " "
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'Psh01:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   Resume Psh02
'Psh02:
'   DoModuleErrors Me
'
'
'End Sub
'
Private Sub cmbRun_Click()
   GetThisRun
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If bCanceled = 0 Then GetThisRun
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4120
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub




Private Sub Form_Activate()
   If bOnLoad Then
      bCanceled = 0
      FillAllRuns cmbPrt
      If optFrom.Value = vbChecked Then
         cmbPrt = ShopSHe02a.cmbPrt
         cmbRun = ShopSHe02a.cmbRun
      End If
      bGoodPart = GetRuns()
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PARUN,RUNREF,RUNSTATUS," _
          & "RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF"

   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter

   bOnLoad = 1
   GetOptions
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   If optFrom Then ShopSHe02a.Show Else FormUnload
   Set ShopSHp01a = Nothing
   
End Sub




Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbRun.Clear
   MouseCursor 13
   sPartNumber = Compress(cmbPrt)
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         If optFrom Then
            cmbRun = ShopSHe02a.cmbRun
         Else
            cmbRun = Format(!Runno, "####0")
         End If
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(!PALEVEL, "#")
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = True
      GetThisRun
   Else
      MouseCursor 0
      sPartNumber = ""
      GetRuns = False
   End If
   MouseCursor 0
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblQty_Click()
   'run qty
   
End Sub



Private Sub optDis_Click()
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbPrt.SetFocus
   Else
      PrintReport
   End If
   
End Sub








Private Sub optPrn_Click()
   PrintReport
   
End Sub



Private Sub GetThisRun()
   Dim RdoRun As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT RUNSTATUS,RUNPKSTART,RUNQTY FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & cmbRun & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblSta = "" & Trim(!RUNSTATUS)
         If Not IsNull(!RUNPKSTART) Then
            sRunPkstart = Format(!RUNPKSTART, "mm/dd/yy")
         Else
            sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
         End If
         lblQty = Format(!RUNQTY, ES_QuantityDataFormat)
         ClearResultSet RdoRun
      End With
   End If
   MouseCursor 0
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
