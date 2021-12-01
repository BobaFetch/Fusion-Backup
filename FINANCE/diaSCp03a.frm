VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaSCp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cost Detail By Part (Report)"
   ClientHeight    =   2580
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2580
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   4080
      Picture         =   "diaSCp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Show BOM Structure"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Tag             =   "3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2580
      FormDesignWidth =   6765
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaSCp03a.frx":0342
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
         Picture         =   "diaSCp03a.frx":04C0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaSCp03a.frx":064A
      PictureDn       =   "diaSCp03a.frx":0790
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   5
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
      PictureUp       =   "diaSCp03a.frx":08D6
      PictureDn       =   "diaSCp03a.frx":0A1C
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For ALL)"
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaSCp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' diaSCp03a - Cost Detail By Part
'
' Notes:
'
' Created: 12/06/02 (nth)
' Revisions:
'   10/22/03 (nth) Added get custom report
'   01/12/03 (nth) Fix report runtime error
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodPart As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbPrt_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbPrt_LostFocus()
   'FindPart Me
End Sub

Private Sub cmdCan_Click()
   bCancel = True
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      If Len(cUR.CurrentPart) Then cmbPrt = cUR.CurrentPart
      FindPart Me
      bOnLoad = False
      FillPartCombo cmbPrt
    '  cmbPrt2 = "ALL"
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodPart = 1 Then
      cUR.CurrentPart = Trim(cmbPrt)
      SaveCurrentSelections
   End If
   FormUnload
   Set diaSCp03a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sPart As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   optPrn.enabled = False
   optDis.enabled = False
   
   If cmbPrt <> "ALL" Then
      sPart = Compress(cmbPrt)
   Else
      sPart = ""
   End If
   
   sCustomReport = GetCustomReport("finsc03.rpt")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   
   sSql = "{PartTable.PARTREF} LIKE '" & sPart & "*' "
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   optPrn.enabled = True
   optDis.enabled = True
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   optPrn.enabled = True
   optDis.enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   Dim sPart As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   optPrn.enabled = False
   optDis.enabled = False
   
   If cmbPrt <> "ALL" Then
      sPart = Compress(cmbPrt)
   Else
      sPart = ""
   End If
   
   'SetMdiReportsize MdiSect
   
   sCustomReport = GetCustomReport("finsc03.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   
   sSql = "{PartTable.PARTREF} LIKE '" & sPart & "*' "
   
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   
   optPrn.enabled = True
   optDis.enabled = True
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   optPrn.enabled = True
   optDis.enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub cmdVew_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then
      ' Part search is closing refresh form
      cmbPrt_LostFocus
   End If
End Sub
