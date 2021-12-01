VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Receipts Register (Report)"
   ClientHeight    =   2685
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2685
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains Customers With Cash Receipts"
      Top             =   720
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaARp04a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaARp04a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   3
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
      PictureUp       =   "diaARp04a.frx":0308
      PictureDn       =   "diaARp04a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5640
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2685
      FormDesignWidth =   6420
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   12
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
      PictureUp       =   "diaARp04a.frx":0594
      PictureDn       =   "diaARp04a.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   3600
      TabIndex        =   14
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   11
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Date"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer "
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "diaARp04a"
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
' Form: diaPar04 - Cash Receipts Register
'
' Notes: Prints or displays a customer invoice register.
'
' Created: (nth)
' Modified:
'   09/25/01 (nth) Fixed errors per WCK and add custom report logic
'   12/04/03 (nth) All CACANCELED to selection formula
'
'************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) Then
      FindCustomer Me, cmbCst
   Else
      cmbCst = "ALL"
   End If
   If cmbCst = "ALL" Then lblNme = "All Customers Selected."
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
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
   Dim rdoCst As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,CACUST FROM " _
          & "CustTable,CashTable WHERE CUREF=CACUST "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         lblNme = ""
         Do Until .EOF
            AddComboStr cmbCst.hwnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoCst = Nothing
   If cmbCst.ListCount > 0 Then
      cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst
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
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   txtEnd = Format(GetServerDateTime, "mm/dd/yy")
   txtBeg = Format(txtEnd, "mm/01/yy")
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Len(cmbCst) And cmbCst <> "ALL" Then
      cUR.CurrentCustomer = cmbCst
      SaveCurrentSelections
   End If
   FormUnload
   Set diaARp04a = Nothing
End Sub

Private Sub PrintReport()
    Dim sCust As String
    Dim sBeg As String
    Dim sEnd As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Title1"
    aFormulaName.Add "Title2"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'For Customer " & CStr(cmbCst) & "'")
    aFormulaValue.Add CStr("'From" & CStr(txtBeg & " Through " & txtEnd) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    sCustomReport = GetCustomReport("finar04.rpt")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
      
   sSql = ""
   sSql = cCRViewer.GetReportSelectionFormula
   
   If (sSql <> "") Then
      sSql = sSql & " AND "
   End If
   
   If UCase(cmbCst) <> "ALL" Then
      sSql = "{CashTable.CACUST} = '" & Compress(cmbCst) & "' AND "
   End If
   
   sSql = sSql & "{CashTable.CARCDATE} >= #" & txtBeg _
          & "# AND {CashTable.CARCDATE} <= #" & txtEnd _
          & "# AND {CashTable.CACANCELED} = 0 "
    
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
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCust As String
   Dim sBeg As String
   Dim sEnd As String
   Dim sCustomReport As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   'SetMdiReportsize MdiSect
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Title1='For Customer " & cmbCst & "'"
   MdiSect.crw.Formulas(2) = "Title2='From " & txtBeg & " Through " & txtEnd & "'"
   MdiSect.crw.Formulas(3) = "RequestBy='Requested By: " & sInitials & "'"
   
   sCustomReport = GetCustomReport("finar04.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = ""
   
   If UCase(cmbCst) <> "ALL" Then
      sSql = "{CashTable.CACUST} = '" & Compress(cmbCst) & "' AND "
   End If
   
   sSql = sSql & "{CashTable.CARCDATE} >= #" & txtBeg _
          & "# AND {CashTable.CARCDATE} <= #" & txtEnd _
          & "# AND {CashTable.CACANCELED} = 0 "
   MdiSect.crw.SelectionFormula = sSql
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
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
