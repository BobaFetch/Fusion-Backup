VERSION 5.00
Begin VB.Form RoutRTp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Labor Efficiency"
   ClientHeight    =   1785
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbRun 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "RoutRTp08a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "RoutRTp08a.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label z1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Run Number"
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PartNumber"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "RoutRTp08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbPrt_LostFocus()
    FillComboRun
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set RoutRTp08a = Nothing
   
End Sub

Private Sub Form_Activate()
   MouseCursor 0
   If bOnLoad = 1 Then FillCombo
   bOnLoad = 0
   MDISect.lblBotPanel = Caption
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())

End Sub
Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "select Distinct PARTNUM from runstable, Parttable " _
          & "where PARTREF = RUNREF ORDER BY PARTNUM"
   LoadComboBox cmbPrt, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Sub FillComboRun()
   On Error GoTo DiaErr1
   sSql = "select Distinct RUNNO from runstable " _
          & "where RUNREF = '" & Compress(cmbPrt) & "' ORDER BY RUNNO"
   LoadComboBox cmbRun, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillRuncombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

'Private Sub PrintReport()
'   MouseCursor 13
'   Dim cCRViewer As EsCrystalRptViewer
'   Dim sCustomReport As String
'   Dim aRptPara As New Collection
'   Dim aRptParaType As New Collection
'   Dim aFormulaValue As New Collection
'   Dim aFormulaName As New Collection
'   Dim sBegDate As String
'   Dim sEndDate As String
'
'
'   Dim sRout As String
'   sRout = Compress(cmbPrt)
'   If Len(sRout) = 0 Then
'      cmbPrt = "ALL"
'      sRout = ""
'   Else
'      If sRout = "ALL" Then sRout = ""
'   End If
'   On Error GoTo DiaErr1
'
'   Set cCRViewer = New EsCrystalRptViewer
'   cCRViewer.Init
'   sCustomReport = GetCustomReport("engrt08")
'   cCRViewer.SetReportFileName sCustomReport, sReportPath
'   cCRViewer.SetReportTitle = sCustomReport
'   aFormulaName.Add "CompanyName"
'   aFormulaName.Add "Includes"
'   aFormulaName.Add "RequestBy"
'
'   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
'   aFormulaValue.Add CStr("'Includes " & CStr(cmbPrt) & "...'")
'   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
'   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
'
'   'sSql = "{TcitTable.TCPARTREF} = '" & Compress(sRout) & "' AND " _
'   '       & " {TcitTable.TCRUNNO} = " & cmbRun
'
'   sSql = "{runstable.runref} = '" & Compress(sRout) & "' AND " _
'          & " {runstable.runno} = " & cmbRun
'
'   cCRViewer.SetReportSelectionFormula sSql
'   cCRViewer.CRViewerSize Me
'   cCRViewer.SetDbTableConnection
'   cCRViewer.ShowGroupTree False
'   cCRViewer.OpenCrystalReportObject Me, aFormulaName
'
'   cCRViewer.ClearFieldCollection aRptPara
'   cCRViewer.ClearFieldCollection aFormulaName
'   cCRViewer.ClearFieldCollection aFormulaValue
'
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'
'End Sub
'

Private Sub PrintReport()
   MouseCursor 13
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sBegDate As String
   Dim sEnddate As String
   
   
   Dim sRout As String
   sRout = Compress(cmbPrt)
'   If Len(sRout) = 0 Then    ' this never worked
'      cmbPrt = "ALL"
'      sRout = ""
'   Else
'      If sRout = "ALL" Then sRout = ""
'   End If
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("engrt08")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = ""   'no longer used. it's now a stored procedure call
          
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   
   ' set parameters
   aRptPara.Add CStr(cmbPrt)
   aRptPara.Add CStr(cmbRun)
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
   
   ' Set report parameter
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
   
   ' run the report
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


