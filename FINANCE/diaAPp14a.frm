VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp14a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Setup (Report)"
   ClientHeight    =   3345
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3345
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkMemo 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaAPp14a.frx":0000
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
         Picture         =   "diaAPp14a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPp14a.frx":0308
      PictureDn       =   "diaAPp14a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3345
      FormDesignWidth =   6480
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPp14a.frx":0594
      PictureDn       =   "diaAPp14a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   3360
      TabIndex        =   13
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Check Number"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1905
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblnme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "diaAPp14a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaAPp14a - view a computer check setup
'
' Notes:
'
' Created: (nth)
' Revisions:
'   10/22/03 (nth) Added getcustomreport
'   08/16/04 (nth) Added saveoptions and getoptions
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


'*************************************************************************************

Private Sub cmbVnd_Click()
   If bCancel = 0 Then
      FindVendor Me
   End If
End Sub

Private Sub cmbVnd_LostFocus()
   If bCancel = 0 Then
      If Trim(cmbVnd) <> "" Then
         FindVendor Me
      Else
         lblNme = "*** Multiple Vendors Selected ***"
      End If
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT CHKVND FROM ChseTable ORDER BY CHKVND"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbVnd.hwnd, "" & Trim(!chkVnd)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCmb = Nothing
   
   If cmbVnd.ListCount > 0 Then
      cmbVnd.ListIndex = 0
   End If
   
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
      If NumberOfChecks = -1 Then
         MouseCursor 0
         'No checks in setup
         MsgBox "Check Setup is Empty.  " _
            & "Please Run Computer Checks Setup.", vbInformation, Caption
         Unload Me
         Exit Sub
      End If
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaAPp14a = Nothing
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finch01.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   If Trim(cmbVnd) <> "" Then
      sSql = "{ChseTable.CHKVND}='" & Trim(cmbVnd) & "'"
      cCRViewer.SetReportSelectionFormula sSql
   End If
   
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
    
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   
   sCustomReport = GetCustomReport("finch01.rpt")
   
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   
   If Trim(cmbVnd) <> "" Then
      MdiSect.crw.SelectionFormula = "{ChseTable.CHKVND}='" & Trim(cmbVnd) & "'"
   End If
   
   'SetCrystalAction Me
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub
