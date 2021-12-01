VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPp16a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cleared Check Summary (Report)"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
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
      FormDesignHeight=   3795
      FormDesignWidth =   7440
   End
   Begin VB.CheckBox ChkExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CheckBox chkComp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.ComboBox txtEndDte 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox txtBegDte 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtEndNum 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtBegNum 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Tag             =   "1"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6240
      TabIndex        =   10
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   19
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
      PictureUp       =   "diaAPp16a.frx":0000
      PictureDn       =   "diaAPp16a.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   20
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
      PictureUp       =   "diaAPp16a.frx":028C
      PictureDn       =   "diaAPp16a.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "External Checks"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Checks"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Check Date"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Check Date"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   14
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Check Number"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   4200
      TabIndex        =   12
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Check Number"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1905
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaAPp16a"
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
'   diaAPp15a - Cleared Check Summary
'
'   Notes:
'
'   Created: 08/07/01 (nth)
'   Revisions:
'       08/16/04 (nth) Added printer to getoptions and saveoptions
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
   bOnLoad = False
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaAPp16a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtBegDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEndDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub PrintReport()
   Dim sBegNum As String
   Dim sEndNum As String
   Dim sBegDte As String
   Dim sEndDte As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   On Error GoTo DiaErr1
   MouseCursor 13
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("finch03.rpt")
   
   sBegNum = Trim(txtBegNum)
   sEndNum = Trim(txtEndNum)
   sBegDte = Trim(txtBegDte)
   sEndDte = Trim(txtEndDte)
   sSql = ""
   
   ' Build selection formula
   If sBegNum <> "" Then
      sSql = "val({ChksTable.CHKNUMBER}) >= " & sBegNum & " AND "
   End If
   If sEndNum <> "" Then
      sSql = sSql & "val({ChksTable.CHKNUMBER}) <= " & sEndNum & " AND "
   End If
   If sBegDte <> "" Then
      sSql = sSql & "{ChksTable.CHKCLEARDATE} >= #" & sBegDte & "# AND "
   End If
   If sEndDte <> "" Then
      sSql = sSql & "{ChksTable.CHKCLEARDATE} <= #" & sEndDte & "# AND "
   End If
   If Not chkComp And chkExt Then
      ' Only external checks
      sSql = sSql & " {ChksTable.CHKTYPE} = 1 AND "
   End If
   If chkComp And Not chkExt Then
      ' Only computer checks
      sSql = sSql & " ISNULL({ChksTable.CHKTYPE}) AND "
   End If
   sSql = sSql & "NOT ISNULL({ChksTable.CHKCLEARDATE})"
   sSql = sSql & " and {ChksTable.CHKTYPE} <> 1   "
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "ChecksNumbered"
    aFormulaName.Add "Checkdates"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Checks Numbered:   " & CStr(sBegNum & "   Through:  " & sEndNum) & "'")
    aFormulaValue.Add CStr("'From:   " & CStr(sBegDte & "   Thru:  " & sEndDte) & "'")
    aFormulaValue.Add CStr("'Requested By ESI'")
    
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
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
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sBegNum As String
   Dim sEndNum As String
   Dim sBegDte As String
   Dim sEndDte As String
   Dim sCustomReport As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   
   'SetMdiReportsize MdiSect
   
   sBegNum = Trim(txtBegNum)
   sEndNum = Trim(txtEndNum)
   sBegDte = Trim(txtBegDte)
   sEndDte = Trim(txtEndDte)
   sSql = ""
   
   ' Build selection formula
   If sBegNum <> "" Then
      sSql = "val({ChksTable.CHKNUMBER}) >= " & sBegNum & " AND "
   End If
   If sEndNum <> "" Then
      sSql = sSql & "val({ChksTable.CHKNUMBER}) <= " & sEndNum & " AND "
   End If
   If sBegDte <> "" Then
      sSql = sSql & "{ChksTable.CHKCLEARDATE} >= #" & sBegDte & "# AND "
   End If
   If sEndDte <> "" Then
      sSql = sSql & "{ChksTable.CHKCLEARDATE} <= #" & sEndDte & "# AND "
   End If
   If Not chkComp And chkExt Then
      ' Only external checks
      sSql = sSql & " {ChksTable.CHKTYPE} = 1 AND "
   End If
   If chkComp And Not chkExt Then
      ' Only computer checks
      sSql = sSql & " ISNULL({ChksTable.CHKTYPE}) AND "
   End If
   sSql = sSql & "NOT ISNULL({ChksTable.CHKCLEARDATE})"
   
   MdiSect.crw.SelectionFormula = sSql
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "ChecksNumbered='Checks Numbered:   " & sBegNum & "   Through:  " & sEndNum & "'"
   MdiSect.crw.Formulas(2) = "Checkdates='From:   " & sBegDte & "   Thru:  " & sEndDte & "'"
   MdiSect.crw.Formulas(3) = "RequestBy='Requested By ESI'"
   
   sCustomReport = GetCustomReport("finch03.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = RTrim(chkComp) & RTrim(chkExt)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(sOptions) > 0 Then
      chkComp = Left(sOptions, 1)
      chkExt = Mid(sOptions, 2, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
