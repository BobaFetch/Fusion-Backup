VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order Document List"
   ClientHeight    =   2715
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Contains Open MO Runs Not Canceled, Complete Or Closed"
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   288
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains MO's Not Canceled"
      Top             =   1080
      Width           =   3345
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Value           =   1  'Checked
      Visible         =   0   'False
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
         Picture         =   "DocuDCp06a.frx":07AE
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
         Picture         =   "DocuDCp06a.frx":092C
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
      FormDesignHeight=   2715
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   9
      Left            =   5400
      TabIndex        =   12
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Part Number"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   10
      Top             =   1440
      Width           =   3072
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5400
      TabIndex        =   8
      Top             =   1440
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Tag             =   " "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1428
   End
End
Attribute VB_Name = "DocuDCp06a"
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

Private Sub FillRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   If cmbPrt = "ALL" Then
      cmbRun = "ALL"
      Exit Sub
   End If
   On Error GoTo DiaErr1
   If lblDsc.ForeColor <> ES_RED Then
      sSql = "SELECT RUNREF,RUNNO,RUNSTATUS FROM RunsTable " _
             & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND " _
             & "RUNSTATUS NOT LIKE 'C%'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
      If bSqlRows Then
         With RdoRns
            Do Until .EOF
               AddComboStr cmbRun.hwnd, str$(!RUNNO)
               .MoveNext
            Loop
            ClearResultSet RdoRns
         End With
      End If
      If cmbRun.ListCount > 0 Then
          cmbRun = cmbRun.List(0)
          If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      End If
   End If
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillCombo()
   Dim sDesc As String
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT RUNREF,PARTREF,PARTNUM FROM " _
          & "RunsTable,PartTable WHERE (RUNREF=PARTREF AND " _
          & "RUNSTATUS<>'CA')"
   LoadComboBox cmbPrt, 1
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   'FillRuns
   'cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   'End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPrt_Click()
   Dim sDesc As String
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   FillRuns
   
End Sub

Private Sub cmbPrt_LostFocus()
   Dim sDesc As String
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   If cmbPrt <> "ALL" Then
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      FillRuns
   Else
      cmbRun.Clear
      cmbRun = "ALL"
      lblDsc = "All Runs Selected"
   End If
   
End Sub

Private Sub cmbRun_LostFocus()
   If Trim(cmbRun) = "" Then cmbPrt = "ALL"
   If cmbPrt = "ALL" Then cmbRun = "ALL"
   
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
   If bOnLoad Then FillCombo
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
   Set DocuDCp06a = Nothing
   
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
   
   On Error GoTo DiaErr1
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDetails"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'MO(s) " & CStr(cmbPrt _
                        & "Run(s) " & cmbRun) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDet.value
   sCustomReport = GetCustomReport("engdc06")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{RunsTable.RUNREF} LIKE '" & sPart & "*' "
   If cmbRun = "ALL" Then
      sSql = sSql & "AND {RunsTable.RUNNO}<99999"
   Else
      sSql = sSql & "AND {RunsTable.RUNNO}=" & Val(cmbRun)
   End If
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


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub GetOptions()
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub
