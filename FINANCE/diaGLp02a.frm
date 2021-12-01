VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Journal Entry List (Report) "
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdcan 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "diaGLp02a.frx":0000
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
         Picture         =   "diaGLp02a.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox txtStart 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2910
      FormDesignWidth =   6120
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp02a.frx":0308
      PictureDn       =   "diaGLp02a.frx":044E
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
      PictureUp       =   "diaGLp02a.frx":0594
      PictureDn       =   "diaGLp02a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include :"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
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
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posted From"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "diaGLp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'************************************************************************************
' diaGLp02a - Journal Entry List (Report)
'
' Notes:
'
' Created: (nth)
'
' Revsions:
'   09/17/03 (nth) Added extended description option button.
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub PrintReport()
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    On Error GoTo DiaErr1
   
    MouseCursor 13
    
    optDis.enabled = False
    optPrn.enabled = False
   
'   SetMdiReportsize MdiSect
   
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy='Requested By: " _
'                        & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "StartDate='" & txtStart & "'"
'   MdiSect.Crw.Formulas(3) = "EndDate='" & txtEnd & "'"
'   MdiSect.Crw.Formulas(4) = "ExtDesc=" & optExt
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "StartDate"
    aFormulaName.Add "EndDate"
    aFormulaName.Add "ExtDesc"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtstart) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtEnd) & "'")
    aFormulaValue.Add optExt
   
    sCustomReport = GetCustomReport("fingl02")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "{GjhdTable.GJPOST} >= #" & txtstart _
          & "# AND {GjhdTable.GJPOST} <= #" & txtEnd & "#"
   
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
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
   
   optDis.enabled = True
   optPrn.enabled = True
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   optDis.enabled = True
   optPrn.enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   
   On Error GoTo DiaErr1
   
   MouseCursor 13
   
   optDis.enabled = False
   optPrn.enabled = False
   
   'SetMdiReportsize MdiSect
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " _
                        & sInitials & "'"
   MdiSect.crw.Formulas(2) = "StartDate='" & txtstart & "'"
   MdiSect.crw.Formulas(3) = "EndDate='" & txtEnd & "'"
   MdiSect.crw.Formulas(4) = "ExtDesc=" & optExt
   
   sCustomReport = GetCustomReport("fingl02")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "{GjhdTable.GJPOST} >= #" & txtstart _
          & "# AND {GjhdTable.GJPOST} <= #" & txtEnd & "#"
   
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   
   optDis.enabled = True
   optPrn.enabled = True
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   optDis.enabled = True
   optPrn.enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

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
   txtstart = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
   GetOptions
   bOnLoad = True
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
   Set diaGLp02a = Nothing
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

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtstart_LostFocus()
   txtstart = CheckDate(txtstart)
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   sOptions = optExt.Value
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optExt.Value = Mid(sOptions, 1, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub
