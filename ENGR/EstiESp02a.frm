VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimate Summary By Customer"
   ClientHeight    =   3645
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   495
      Left            =   2040
      TabIndex        =   26
      Top             =   2100
      Width           =   4095
      Begin VB.OptionButton optAcc 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   200
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optAcc 
         Caption         =   "Accepted"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   200
         Width           =   1215
      End
      Begin VB.OptionButton optAcc 
         Caption         =   "Not Accepted"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   6
         Top             =   200
         Width           =   1575
      End
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "EstiESp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESp02a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   2880
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   3120
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1800
      Width           =   1250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1800
      Width           =   1250
   End
   Begin VB.TextBox txtCst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "EstiESp02a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "EstiESp02a.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Estimates"
      Top             =   1080
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3645
      FormDesignWidth =   7215
   End
   Begin VB.Label lblNotice 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   5772
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   252
      Index           =   8
      Left            =   5520
      TabIndex        =   22
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   252
      Index           =   4
      Left            =   5520
      TabIndex        =   21
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Estimates"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Desc"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Tag             =   " "
      Top             =   3120
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Tag             =   " "
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   16
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimates From"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2025
   End
End
Attribute VB_Name = "EstiESp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Click()
   GetTheCustomer
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(Trim(cmbCst)) = 0 Then cmbCst = cmbCst.List(0)
   GetTheCustomer
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BIDCUST,CUREF,CUNICKNAME " _
          & "FROM EstiTable,CustTable WHERE BIDCUST=CUREF " _
          & "ORDER BY CUREF"
   LoadComboBox cmbCst, 1
   cmbCst = "ALL"
   txtCst = "All Customers Selected."
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set EstiESp02a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBDate As String
   Dim sEDate As String
   Dim sCust As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If Not IsDate(txtBeg) Then
      sBDate = "1995,01,01"
   Else
      sBDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "2024,12,31"
   Else
      sEDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   ' MM TODO SetMdiReportsize MDISect
   
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   If Trim(cmbCst) <> "ALL" Then sCust = Compress(cmbCst)

   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDescription"
   aFormulaName.Add "ShowExDescription"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s) " & CStr(cmbCst & " From " _
                       & txtBeg & " Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.value
   aFormulaValue.Add optExt.value
   
   sCustomReport = GetCustomReport("enges02")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{EstiTable.BIDCUST} LIKE '" & sCust & "*' " _
          & "AND ({EstiTable.BIDDATE} In Date(" & sBDate & ") " _
          & "To Date(" & sEDate & "))"
   
   If optAcc(1).value = True Then
      sSql = sSql & " AND {EstiTable.BIDACCEPTED}=1"
   Else
      If optAcc(2).value = True Then sSql = sSql & " AND {EstiTable.BIDACCEPTED}=0"
   End If
   sSql = sSql & " AND {EstiTable.BIDCANCELED} = 0"
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
'   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   txtEnd = "ALL"
   txtBeg = "ALL"
   txtCst.BackColor = Es_FormBackColor

'removed R63
'   If UCase$(Left(MdiSect.Caption, 4)) = "SALE" Then
'      optAcc(1).Value = True
'      z2.Enabled = False
'      lblNotice.Caption = "Only Accepted Estimates Are Reported In The Sales Section"
'   End If
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = RTrim(optDsc.value) _
              & RTrim(optExt.value)
   SaveSetting "Esi2000", "EsiEngr", "es02", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "Pes02", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiEngr", "es02", sOptions)
   If Len(sOptions) > 0 Then
      optDsc.value = Val(Left(sOptions, 1))
      optExt.value = Val(Mid(sOptions, 2, 1))
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pes02", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub



Private Sub GetTheCustomer()
   Dim RdoCst As ADODB.Recordset
   sSql = "Qry_GetCustomerBasics '" & Compress(cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         txtCst = (!CUNAME)
         ClearResultSet RdoCst
      End With
   Else
      If cmbCst = "ALL" Then
         txtCst = "All Customers Selected."
      Else
         txtCst = "Range Of Customers Selected."
      End If
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthecust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub
