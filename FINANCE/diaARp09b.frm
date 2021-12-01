VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARp09b 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Invoiced and Uninvoiced Cash Receipts"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6855
   Begin VB.CheckBox chkIncludeNonInvoiced 
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   2940
      Width           =   435
   End
   Begin VB.ComboBox cmbChkNum 
      Height          =   315
      ItemData        =   "diaARp09b.frx":0000
      Left            =   1800
      List            =   "diaARp09b.frx":0002
      TabIndex        =   17
      Tag             =   "3"
      Top             =   2400
      Width           =   1320
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   600
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaARp09b.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaARp09b.frx":0182
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      Top             =   465
      Width           =   1555
   End
   Begin VB.ComboBox txtstart 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1920
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   13
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
      PictureUp       =   "diaARp09b.frx":030C
      PictureDn       =   "diaARp09b.frx":0452
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   14
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
      PictureUp       =   "diaARp09b.frx":0598
      PictureDn       =   "diaARp09b.frx":06DE
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3585
      FormDesignWidth =   6855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include non-invoiced receipts"
      Height          =   165
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   2355
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   165
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   2404
      Width           =   1335
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank If Unknown)"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amount"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank If Unknown)"
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "diaARp09b"
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
'
' diaAPe09a - View as cash receipts
'
' Created: 09/01/03 (JCW)
' Revisions:
'   09/04/03 (nth) Added to esifina
'   02/11/04 (nth) Added compress customer name to the print report function
'                  selection query fix.
'   06/24/04 (nth) Added optional parameters to printreport so report can be queue
'                  from anothor form
'   08/16/04 (nth) Added getoptions and saveoptions
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim sMsg As String
Public bRemote As Byte

' Key Handeling
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


'*************************************************************************************

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   FindCustomer Me, cmbCst
   ' Get all Check number for the customer
   GetCheckNumber
   
End Sub

Private Sub Form_Load()
   GetOptions
   If bRemote Then
      Me.WindowState = vbMinimized
   Else
      FormLoad Me
      FormatControls
      bOnLoad = True
   End If
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
End Sub

Public Sub FillCombo()
   On Error GoTo DiaErr1
   FillCustomers Me
   FindCustomer Me, cmbCst
   ' Get all Check number for the customer
   GetCheckNumber
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   SaveOptions
   If Not bRemote Then
      FormUnload
   End If
   bRemote = False
   Set diaARp09b = Nothing
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub optDis_Click()
   If Not bRemote Then
      PrintReport
   End If
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtAmt_Change()
   ' Get all Check number for the customer
   GetCheckNumber
End Sub

Private Sub txtAmt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtAmt_LostFocus()
   txtAmt = Format(txtAmt, CURRENCYMASK)
End Sub

Private Sub txtStart_Change()
   ' Get all Check number for the customer
   'GetCheckNumber
End Sub

Private Sub txtstart_LostFocus()
   If Trim(txtstart) <> "" Then
      txtstart = CheckDate(txtstart)
      GetCheckNumber
   End If
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtstart_GotFocus()
   SelectFormat Me
End Sub

Public Sub PrintReport(Optional sCust As String, Optional sNum As String)
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   If sCust <> "" Then
      cmbCst = sCust
   End If
   If Trim(cmbCst) = "" Then
      sMsg = "Please Select A Customer."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MdiSect.Crw.Formulas(2) = "Title1 = 'From: " & Trim(txtstart) & "'"
'   MdiSect.Crw.Formulas(3) = "Title2 = '" & Trim(txtAmt) & "'"
'   MdiSect.Crw.Formulas(4) = "Title3 = 'Customer: " & Trim(cmbCst) & "'"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title2"
   aFormulaName.Add "Title3"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'From: " & CStr(Trim(txtstart)) & "'")
   aFormulaValue.Add CStr("'" & CStr(Trim(txtAmt)) & "'")
   aFormulaValue.Add CStr("'Customer: " & CStr(Trim(cmbCst)) & "'")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("finar09b.rpt")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   
   'sSql = "{JritTable.DCCREDIT} <> 0 and {CashTable.CACUST} = '" _
   '    & Compress(cmbCst.Text) & "'"
   'removed jrittable dependency 2/1/06
'   sSql = "{CashTable.CACUST} = '" _
'          & Compress(cmbCst.Text) & "'"
'   If sNum <> "" Then
'      bRemote = True
'      optDis = True
'      sSql = sSql & " AND {CashTable.CACHECKNO} = '" & sNum & "'"
'   Else
'      If Trim(txtstart) <> "" Then
'         sSql = sSql & " and {CashTable.CARCDATE} >= DateTime ('" _
'                & Trim(txtstart) & "')"
'      End If
'      If Trim(txtAmt) <> "" Then
'         sSql = sSql & " AND ccur({CashTable.CACKAMT}) = ccur('" _
'                & Trim(txtAmt.Text) & "')"
'      End If
'   End If
'
'   If Trim(cmbChkNum) <> "<ALL>" Then
'      sSql = sSql & " and {CashTable.CACHECKNO} = '" & Trim(cmbChkNum) & "'"
'   End If
'
'    ' added in v 18.0.6
'    If chkIncludeNonInvoiced.Value = vbUnchecked Then
'       sSql = sSql & " AND not(isnull({CihdTable.INVDATE}))"
'    End If
'
'   cCRViewer.SetReportSelectionFormula sSql
   
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("String")
   aRptParaType.Add CStr("Int")
   
   aRptPara.Add CStr(cmbCst)                    '@CustNickName
   aRptPara.Add CStr(txtstart)                  '@StartDate
   aRptPara.Add CStr(txtAmt)                    '@ReceiptAmount
   aRptPara.Add CStr(cmbChkNum)                 '@CheckNumber
   aRptPara.Add chkIncludeNonInvoiced.Value     '@ShowUninvoicedItems
   
   ' Set report parameter

   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
  
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

'Public Sub PrintReport1(Optional sCust As String, Optional sNum As String)
'   On Error GoTo DiaErr1
'   If sCust <> "" Then
'      cmbCst = sCust
'   End If
'   If Trim(cmbCst) = "" Then
'      sMsg = "Please Select A Customer."
'      MsgBox sMsg, vbInformation, Caption
'      Exit Sub
'   End If
'
'   MouseCursor 13
'
'   'SetMdiReportsize MdiSect
'   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MdiSect.crw.Formulas(2) = "Title1 = 'From: " & Trim(txtstart) & "'"
'   MdiSect.crw.Formulas(3) = "Title2 = '" & Trim(txtAmt) & "'"
'   MdiSect.crw.Formulas(4) = "Title3 = 'Customer: " & Trim(cmbCst) & "'"
'
'   sCustomReport = GetCustomReport("finar09b.rpt")
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
'
'   'sSql = "{JritTable.DCCREDIT} <> 0 and {CashTable.CACUST} = '" _
'   '    & Compress(cmbCst.Text) & "'"
'   'removed jrittable dependency 2/1/06
'   sSql = "{CashTable.CACUST} = '" _
'          & Compress(cmbCst.Text) & "'"
'   If sNum <> "" Then
'      bRemote = True
'      optDis = True
'      sSql = sSql & " AND {CashTable.CACHECKNO} = '" & sNum & "'"
'   Else
'      If Trim(txtstart) <> "" Then
'         sSql = sSql & " and {CashTable.CARCDATE} >= DateTime ('" _
'                & Trim(txtstart) & "')"
'      End If
'      If Trim(txtAmt) <> "" Then
'         sSql = sSql & " AND ccur({CashTable.CACKAMT}) = ccur('" _
'                & Trim(txtAmt.Text) & "')"
'      End If
'   End If
'
'   If Trim(cmbChkNum) <> "<ALL>" Then
'      sSql = sSql & " and {CashTable.CACHECKNO} = '" & Trim(cmbChkNum) & "'"
'   End If
'
'
'   MdiSect.crw.SelectionFormula = sSql
'   'SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'DiaErr1:
'   sProcName = "PrintReport"
'   CurrError.Number = Err
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub
'
Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub GetCheckNumber()

    Dim sCustomer As String
    Dim sAmt As String
    Dim sStartDate As String
    Dim sChkSQL As String
    
    If Len(Trim(cmbCst.Text)) = 0 Then Exit Sub ' At least one customer name should be present
    sCustomer = cmbCst.Text
    
    sStartDate = Trim(txtstart.Text)
    'sAmt = Trim(txtAmt.Text)
        sAmt = Replace(Trim(txtAmt.Text), Chr$(44), "")
    
    sChkSQL = "Select DISTINCT CACHECKNO from CashTable " & vbCrLf _
          & " WHERE CACUST = '" & sCustomer & "'"
    
    
    If sStartDate <> "" Then
        sChkSQL = sChkSQL & " AND CARCDATE = '" & sStartDate & "'"
    End If
    
    If sAmt <> "" Then
        sChkSQL = sChkSQL & " AND CACKAMT = '" & sAmt & "'"
    End If
    
    ' Load the Check combo box
    LoadComboWithSQL cmbChkNum, sChkSQL, True
    Debug.Print sChkSQL
    
End Sub


