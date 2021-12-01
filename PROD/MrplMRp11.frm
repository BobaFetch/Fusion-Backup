VERSION 5.00
Begin VB.Form MrplMRp11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculate ROLT by Class & Product Code"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkPartType 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   25
      Top             =   1680
      Width           =   615
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   24
      Top             =   1680
      Width           =   615
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   23
      Top             =   1680
      Width           =   615
   End
   Begin VB.CheckBox chkPartType 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   22
      Top             =   1680
      Width           =   615
   End
   Begin VB.CheckBox chkExactMatch 
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   2160
      Width           =   495
   End
   Begin VB.CheckBox chkShowDetails 
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtPartPrefix 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Tag             =   "3"
      Top             =   240
      Width           =   2895
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Select Product Class From List"
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Select Product Code From List"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5880
      Picture         =   "MrplMRp11.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Find Part"
      Top             =   240
      Visible         =   0   'False
      Width           =   395
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Close"
      Height          =   360
      Left            =   8400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Requires A Valid Part Number"
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Include Part Types:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Exact Match Only?"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Include Components?"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblRecords 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Index           =   6
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Part Class Prefix"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code Prefix"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number Prefix"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "MrplMRp11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

'Reorder Lead Time by Clss and Product Code Report
'7/8/2015 TEL New Report

Option Explicit

Dim bOnLoad As Byte
Dim bAtLeastOneDefaultRouting As Byte
'Dim cAvgWorkWeekHrs As Currency

'Dim iWorkDays As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    ViewParts.lblControl = "txtPartPrefix"
    ViewParts.txtPrt = txtPartPrefix
    ViewParts.Show
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
      'If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCde.AddItem "", 0
      
      FillProductClasses
      'If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      cmbCls.AddItem "", 0
   End If
   bOnLoad = 0
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   'SetupReportTables
   'cAvgWorkWeekHrs = GetAvgWorkWeekHrs
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
    Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set MrplMRp07a = Nothing
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim I As Integer
   
   On Error Resume Next
       
   sOptions = Left(Trim(txtPartPrefix) & Space(30), 30)
   sOptions = sOptions & Left(Trim(cmbCde) & Space(6), 6)
   sOptions = sOptions & Left(Trim(cmbCls) & Space(4), 4)
   sOptions = sOptions & Left(CStr(chkExactMatch.Value) & Space(6), 6)
   sOptions = sOptions & Left(CStr(chkShowDetails.Value) & Space(6), 6)
   sOptions = sOptions & Left(CStr(chkPartType(1).Value) & Space(6), 6)
   sOptions = sOptions & Left(CStr(chkPartType(2).Value) & Space(6), 6)
   sOptions = sOptions & Left(CStr(chkPartType(3).Value) & Space(6), 6)
   sOptions = sOptions & Left(CStr(chkPartType(4).Value) & Space(6), 6)
   
   SaveSetting "Esi2000", "EsiProd", "mrp11", Trim(sOptions)
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim I As Integer

   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "mrp11", Trim(sOptions))
   If Len(sOptions) > 0 Then txtPartPrefix = Trim(Mid(sOptions, 1, 30)) Else txtPartPrefix = ""
   If Len(sOptions) > 30 Then cmbCde = Trim(Mid(sOptions, 31, 6)) Else cmbCde = ""
   If Len(sOptions) > 36 Then cmbCls = Trim(Mid(sOptions, 37, 4)) Else cmbCls = ""
   If Len(sOptions) > 40 Then chkExactMatch.Value = Trim(Mid(sOptions, 41, 6)) Else chkExactMatch.Value = False
   If Len(sOptions) > 46 Then chkShowDetails.Value = Trim(Mid(sOptions, 47, 6)) Else chkShowDetails = False
   If Len(sOptions) > 52 Then chkPartType(1).Value = Trim(Mid(sOptions, 53, 6)) Else chkPartType(1).Value = False
   If Len(sOptions) > 58 Then chkPartType(2).Value = Trim(Mid(sOptions, 59, 6)) Else chkPartType(2).Value = False
   If Len(sOptions) > 64 Then chkPartType(3).Value = Trim(Mid(sOptions, 65, 6)) Else chkPartType(3).Value = False
   If Len(sOptions) > 70 Then chkPartType(4).Value = Trim(Mid(sOptions, 71, 6)) Else chkPartType(4).Value = False
End Sub



Private Sub RemoveReportData()
    'Remove all data from table
    sSql = "DELETE FROM EsReportROLT WHERE ROLTUser = '" & sInitials & "'"
    clsADOCon.ExecuteSql sSql
    
    
    'Remove all data from table
    sSql = "DELETE FROM EsReportROLTDetail WHERE ROLTUser = '" & sInitials & "'"
    clsADOCon.ExecuteSql sSql

End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
    PrintReport
End Sub



Private Sub PrintReport()

   ' ask before generating report for all parts
   If Compress(txtPartPrefix.Text) = "" And Trim(cmbCde) = "" And Trim(cmbCls) = "" Then
      If MsgBox("Generate Report for all parts of selected types?", vbQuestion + vbYesNo, "Generate for all parts?") <> vbYes Then
         Return
      End If
   End If
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   Dim sPart As String
   Dim sClass As String
   
   sCustomReport = GetCustomReport("prdmr11")
    
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath

   cCRViewer.SetReportTitle = "prdmr11"
   cCRViewer.ShowGroupTree False
    
   cCRViewer.CRViewerSize Me
   
   cCRViewer.SetDbTableConnection True    'True if stored procedure call
   
   'add sp report parameters after SetDbTableConnection call because it clears them
   aRptPara.Add Compress(txtPartPrefix.Text)
   aRptPara.Add Compress(cmbCde.Text)
   aRptPara.Add Compress(cmbCls.Text)
   aRptPara.Add chkShowDetails.Value
   aRptPara.Add chkExactMatch.Value
   aRptPara.Add chkPartType(1).Value
   aRptPara.Add chkPartType(2).Value
   aRptPara.Add chkPartType(3).Value
   aRptPara.Add chkPartType(4).Value
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Bool")
   aRptParaType.Add CStr("Bool")
   aRptParaType.Add CStr("Bool")
   aRptParaType.Add CStr("Bool")
   aRptParaType.Add CStr("Bool")
   aRptParaType.Add CStr("Bool")
   
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType   'must happen AFTER SetDbTableConnection call!
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

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      txtPartPrefix.Visible = False
      cmdFind.Visible = True
   Else
      txtPartPrefix.Visible = True
      cmdFind.Visible = False
   End If
End Function

