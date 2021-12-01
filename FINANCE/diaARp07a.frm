VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp07a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales By GL Account (Report)"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6810
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5640
      TabIndex        =   9
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaARp07a.frx":0000
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
         Picture         =   "diaARp07a.frx":018A
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
      Left            =   5640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "4"
      ToolTipText     =   "Contains Customers With Aging"
      Top             =   600
      Width           =   1080
   End
   Begin VB.CheckBox chkDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CheckBox chkSum 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3375
      Width           =   975
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Tag             =   "2"
      Top             =   2280
      Width           =   1555
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
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
      PictureUp       =   "diaARp07a.frx":0308
      PictureDn       =   "diaARp07a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3720
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4005
      FormDesignWidth =   6810
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   11
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
      PictureUp       =   "diaARp07a.frx":0594
      PictureDn       =   "diaARp07a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   4200
      TabIndex        =   22
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   21
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Sales From"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Part Description?"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   14
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Report Only?"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1305
   End
End
Attribute VB_Name = "diaARp07a"
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
' diaAPp10a - Sales By GL Account
'
' Notes:
'
' Created: 10/10/03 (JCW)
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbAct_Click()
   FindAccount Me
End Sub

Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If Trim(cmbAct) = "" Then
      lblDsc = "Muliple Accounts Selected."
   Else
      FindAccount Me
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Trim(cmbCst) = "" Then
      lblNme = "Multiple Customers Selected."
   Else
      FindCustomer Me, cmbCst
   End If
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
      FillCombo
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   sCurrForm = Caption
   bOnLoad = True
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set diaAPp10a = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   '    SysPrinters.Show
   '    ShowPrinters.Value = False
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
   
   MouseCursor 13
   On Error GoTo DiaErr1
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "... '")
    aFormulaValue.Add CStr("'Sales By GL Account'")
    aFormulaValue.Add CStr("'From " & CStr(txtBeg & " Through " & txtEnd) & "... '")
   
   
    sCustomReport = GetCustomReport("finar07.rpt")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport

'{JritTable.DCCREDIT} <> 0 and
   sSql = "{JrhdTable.MJTYPE} = 'sj'" _
          & " and {CihdTable.INVDATE} >= DateTime ('" & Trim(txtBeg) & "') and " _
          & "{CihdTable.INVDATE} <= DateTime ('" & Trim(txtEnd) & "')"
   If Trim(cmbCst) <> "" Then
       aFormulaName.Add "Customer"
       aFormulaValue.Add CStr("'Customer: " & CStr(cmbCst) & "'")
      sSql = sSql & " and {JritTable.DCCUST} = '" & Compress(cmbCst) & "' "
       aFormulaName.Add "CompressCust"
       aFormulaValue.Add CStr("'Customer: " & CStr(cmbCst) & "'")
   Else
       aFormulaName.Add "Customer"
       aFormulaValue.Add CStr("'Customers: All'")
       aFormulaName.Add "CompressCust"
       aFormulaValue.Add CStr("'All'")
   End If
   If Trim(cmbAct) <> "" Then
      sSql = sSql & " and {GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "' "
      aFormulaName.Add "incAccts"
      aFormulaValue.Add CStr("'Account: " & CStr(Trim(cmbAct)) & "'")
      aFormulaName.Add "account"
      aFormulaValue.Add CStr("'Account: " & CStr(Compress(cmbAct)) & "'")
   Else
      aFormulaName.Add "incAccts"
      aFormulaValue.Add CStr("'Accounts: All'")
      aFormulaName.Add "account"
      aFormulaValue.Add CStr("'All'")
   End If
   If chkDsc.Value = vbUnchecked Then
      aFormulaName.Add "Dsc"
      aFormulaValue.Add CStr("'0'")
   Else
      aFormulaName.Add "Dsc"
      aFormulaValue.Add CStr("'1'")
   End If
   If chkSum.Value = vbChecked Then
      aFormulaName.Add "Sum"
      aFormulaValue.Add CStr("'1'")
   Else
      aFormulaName.Add "Sum"
      aFormulaValue.Add CStr("'0'")
   End If
      aFormulaName.Add "BegDate"
      aFormulaValue.Add CStr("'" & CStr(Trim(txtBeg)) & "'")
      aFormulaName.Add "EndDate"
      aFormulaValue.Add CStr("'" & CStr(Trim(txtEnd)) & "'")
   
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
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomReport As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='Sales By GL Account'"
   MdiSect.crw.Formulas(3) = "Title2='From " & txtBeg & " Through " & txtEnd & "'"
   
   
   sCustomReport = GetCustomReport("finar07.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = "{JritTable.DCCREDIT} <> 0 and {JrhdTable.MJTYPE} = 'sj'" _
          & " and {CihdTable.INVDATE} >= DateTime ('" & Trim(txtBeg) & "') and " _
          & "{CihdTable.INVDATE} <= DateTime ('" & Trim(txtEnd) & "')"
   If Trim(cmbCst) <> "" Then
      MdiSect.crw.Formulas(4) = "Customer='Customer: " & cmbCst & "'"
      sSql = sSql & " and {JritTable.DCCUST} = '" & Compress(cmbCst) & "' "
      MdiSect.crw.Formulas(5) = "CompressCust='" & Compress(cmbCst) & "'"
   Else
      MdiSect.crw.Formulas(4) = "Customer='Customers: All'"
      MdiSect.crw.Formulas(5) = "CompressCust='All'"
   End If
   If Trim(cmbAct) <> "" Then
      sSql = sSql & " and {GlacTable.GLACCTREF} = '" & Compress(cmbAct) & "' "
      MdiSect.crw.Formulas(6) = "incAccts='Account: " & Trim(cmbAct) & "'"
      MdiSect.crw.Formulas(7) = "account='" & Compress(cmbAct) & "'"
   Else
      MdiSect.crw.Formulas(6) = "incAccts='Accounts: All'"
      MdiSect.crw.Formulas(7) = "account='All'"
   End If
   If chkDsc.Value = vbUnchecked Then
      MdiSect.crw.Formulas(8) = "Dsc = '0'"
   Else
      MdiSect.crw.Formulas(8) = "Dsc = '1'"
   End If
   If chkSum.Value = vbChecked Then
      MdiSect.crw.Formulas(9) = "Sum = '1'"
   Else
      MdiSect.crw.Formulas(9) = "Sum = '0'"
   End If
   
   MdiSect.crw.Formulas(10) = "BegDate='" & Trim(txtBeg) & "'"
   MdiSect.crw.Formulas(11) = "EndDate='" & Trim(txtEnd) & "'"
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillCombo()
   Dim rdoCombo As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccountCombo"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCombo, ES_FORWARD)
   If bSqlRows Then
      With rdoCombo
         Do While Not .EOF
            AddComboStr cmbAct.hwnd, Trim(!GLACCTNO)
            .MoveNext
         Loop
         cmbAct.ListIndex = 0
      End With
   End If
   Set rdoCombo = Nothing
   FillCustomers Me
   FindCustomer Me, cmbCst
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Public Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(chkDsc.Value) _
              & RTrim(chkSum.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      chkDsc.Value = Val(Mid(sOptions, 1, 1))
      chkSum.Value = Val(Mid(sOptions, 2, 1))
   Else
      chkDsc.Value = vbChecked
      chkSum.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
