VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaGLp16a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Balance"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaGLp16a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaGLp16a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5100
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2340
      FormDesignWidth =   6720
   End
   Begin VB.ComboBox cboAccount 
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Tag             =   "3"
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox cboEnd 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1860
      Width           =   1095
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1500
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   6
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
      PictureUp       =   "diaGLp16a.frx":0308
      PictureDn       =   "diaGLp16a.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   10
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
      PictureUp       =   "diaGLp16a.frx":0594
      PictureDn       =   "diaGLp16a.frx":06DA
   End
   Begin VB.Label lblAccount 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   1140
      Width           =   2775
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Balance for this Account"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show balance as of this date"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1860
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail Starting"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Top             =   1500
      Width           =   2655
   End
End
Attribute VB_Name = "diaGLp16a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

' See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' diaGLp16a - Account Balance (Report)
'
' Notes: Same form used for both reports.
'
' Created: 03/5/06 (TEL)

'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

' bForm documentation
' 0 = Detail GL
' 1 = Trial Balance
Dim bForm As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub PrintReport()
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
    On Error GoTo DiaErr1
    'SetMdiReportsize MdiSect
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("fingl16.rpt")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Start"
    aFormulaName.Add "End"
    aFormulaName.Add "Account"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboStart) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboEnd) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboAccount) & "'")
    aFormulaValue.Add CStr("'Account Balance as of " & CStr(Format(cboEnd, "m/d/yy")) & "'")
    aFormulaValue.Add CStr("'Account: " & CStr(cboAccount) & "'")
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    sSql = "{GlActivityView.Account} = {@Account} and " _
           & " {GlActivityView.XtnDate} in {@StartDate} to {@EndDate}"
    'cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   
   sCustomReport = GetCustomReport("fingl16.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Start='" & cboStart & "'"
   MdiSect.crw.Formulas(3) = "End='" & cboEnd & "'"
   MdiSect.crw.Formulas(4) = "Account='" & cboAccount & "'"
   
   MdiSect.crw.Formulas(5) = "Title1='Account Balance as of " _
                        & Format(cboEnd, "m/d/yy") & "'"
   
   MdiSect.crw.Formulas(6) = "Title2='Account: " & cboAccount & "'"
   
   'SetCrystalAction Me
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cboAccount_Click()
   Dim gl As New GLTransaction
   lblAccount = gl.GetAccountName(cboAccount.Text)
End Sub

Private Sub cmbAccount_LostFocus()
   Dim gl As New GLTransaction
   lblAccount = gl.GetAccountName(cboAccount.Text)
End Sub

Private Sub cboEnd_Click()
   On Error Resume Next
'   cboStart = CheckDate(DateAdd("d", -Day(cboEnd) + 1, cboEnd.Text))
End Sub

Private Sub cboEnd_GotFocus()
   On Error Resume Next
   'cboStart = CheckDate(DateAdd("d", -Day(cboEnd) + 1, cboEnd.Text))
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
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
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   Dim gl As New GLTransaction
   gl.FillComboWithAccounts cboAccount
   lblAccount = gl.GetAccountName(cboAccount.Text)
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
   FormUnload
   On Error Resume Next
   Set diaGLp16a = Nothing
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

Private Sub cboEnd_DropDown()
   On Error Resume Next
   ShowCalendar Me
   'cboStart = DateAdd("d", -Day(cboEnd) + 1, cboEnd.Text)
End Sub

Private Sub cboEnd_LostFocus()
   cboEnd = CheckDate(cboEnd)
   'cboStart = CheckDate(DateAdd("d", -Day(cboEnd) + 1, cboEnd.Text))
End Sub

Private Sub cboStart_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboStart_LostFocus()
   cboStart = CheckDate(cboStart)
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & "StartDate", cboStart
   SaveSetting "Esi2000", "EsiFina", Me.Name & "EndDate", cboEnd
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Account", cboAccount
End Sub

Private Sub GetOptions()
   Dim defaultDate As String
   defaultDate = Format(Date, "mm/dd/yyyy")
   cboStart = GetSetting("Esi2000", "EsiFina", Me.Name & "StartDate", defaultDate)
   cboEnd = GetSetting("Esi2000", "EsiFina", Me.Name & "EndDate", defaultDate)
   cboAccount = GetSetting("Esi2000", "EsiFina", Me.Name & "Account", cboAccount.List(0))
   Dim gl As New GLTransaction
   lblAccount = gl.GetAccountName(cboAccount.Text)
   
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub
