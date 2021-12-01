VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Allocations By Customer"
   ClientHeight    =   3660
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3660
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "ShopSHp11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
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
      Picture         =   "ShopSHp11a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtSon 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Tag             =   "1"
      Text            =   "000000"
      ToolTipText     =   "Sales Order Number (No Class)"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox optSta 
      Caption         =   "SC"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Scheduled"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "RL"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      ToolTipText     =   "Pick List"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PL"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Picked Partial"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PP"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      ToolTipText     =   "Picked Complete"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PC"
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "Complete"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CO"
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Closed"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CL"
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   9
      ToolTipText     =   "Canceled"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CA"
      Height          =   255
      Index           =   7
      Left            =   6360
      TabIndex        =   10
      ToolTipText     =   "Released"
      Top             =   2280
      Width           =   615
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1800
      Width           =   1250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1800
      Width           =   1250
   End
   Begin VB.TextBox txtNme 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp11a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp11a.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   1080
      Width           =   1555
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   3360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3660
      FormDesignWidth =   7230
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   29
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Order Allocations To:"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   27
      Top             =   720
      Width           =   2865
   End
   Begin VB.Label lblSoType 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   26
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All - 00000)"
      Height          =   285
      Index           =   8
      Left            =   5520
      TabIndex        =   25
      Top             =   2760
      Width           =   1700
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Number"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Sales Order Number (No Class)"
      Top             =   2640
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Status:"
      Height          =   15
      Index           =   5
      Left            =   0
      TabIndex        =   23
      Top             =   990
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Status:"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   2280
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   21
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   1908
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   5520
      TabIndex        =   19
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Sched Compl From"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   2145
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   17
      Top             =   1800
      Width           =   915
   End
End
Attribute VB_Name = "ShopSHp11a"
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
   GetReportCustomer
   txtSon = SO_NUM_FORMAT
   lblSoType = ""
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   GetReportCustomer
   
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
   sSql = "Qry_FillCustomerCombo"
   LoadComboBox cmbCst, -1
   cmbCst = "ALL"
   GetReportCustomer
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
   Set ShopSHp11a = Nothing
   
End Sub
Private Sub PrintReport()
   Dim bByte As Byte
   Dim lSon As Long
   Dim sBDate As String
   Dim sEDate As String
   Dim sCust As String
   Dim sStatus As String
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   bByte = GetAllocations
   If bByte = 0 Then
      MsgBox "There Are No Manufacturing Order Allocations.", _
         vbInformation, Caption
      Exit Sub
   End If
   MouseCursor 13
   On Error GoTo DiaErr1
   lSon = Val(txtSon)
   If cmbCst = "ALL" Then
      sCust = ""
   Else
      sCust = Compress(cmbCst)
   End If
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
   sCustomReport = GetCustomReport("prdsh18")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & txtNme & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
  ' MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{CustTable.CUREF} LIKE '" & sCust & "*' " _
          & "AND ({RunsTable.RUNSCHED} In Date(" & sBDate & ") " _
          & "To Date(" & sEDate & ")) "
   If optSta(0).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'SC' "
   If optSta(1).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'RL' "
   If optSta(2).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PL' "
   If optSta(3).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PP' "
   If optSta(4).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PC' "
   If optSta(5).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'C0' "
   If optSta(6).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CL' "
   If optSta(7).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CA' "
   If lSon > 0 Then sSql = sSql & "AND {SohdTable.SONUMBER}=" & lSon & " "
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me

   cCRViewer.SetReportSelectionFormula sSql
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
   txtEnd = ""
   txtBeg = ""
   txtSon = SO_NUM_FORMAT
   txtNme.BackColor = Es_FormBackColor
   
End Sub

Private Sub SaveOptions()
   Dim b As Byte
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiProd", "sh18Printer", lblPrinter
   
   For b = 0 To 6
      sOptions = sOptions & Trim(str$(optSta(b).Value))
   Next
   sOptions = sOptions & Trim(str$(optSta(b).Value))
   SaveSetting "Esi2000", "EsiProd", "sh18a", sOptions
   
   
End Sub

Private Sub GetOptions()
   Dim b As Byte
   Dim sOptions As String
   
   On Error Resume Next
   lblPrinter = GetSetting("Esi2000", "EsiProd", "sh18Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   sOptions = GetSetting("Esi2000", "EsiProd", "sh18a", sOptions)
   If Len(sOptions) = 8 Then
      For b = 0 To 6
         optSta(b).Value = Val(Mid$(sOptions, b + 1, 1))
      Next
      optSta(b).Value = Val(Mid$(sOptions, b + 1, 1))
   Else
      For b = 0 To 6
         optSta(b).Value = vbChecked
      Next
      optSta(b).Value = vbChecked
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub



Private Sub GetReportCustomer()
   Dim RdoCst As ADODB.Recordset
   
   On Error GoTo DiaErr1
   If Trim(cmbCst) <> "ALL" Then
      sSql = "Qry_GetCustomerBasics '" & Compress(cmbCst) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
      If bSqlRows Then
         With RdoCst
            cmbCst = "" & Trim(!CUNICKNAME)
            txtNme = "" & Trim(!CUNAME)
            ClearResultSet RdoCst
         End With
      Else
         txtNme = "*** Multiple Customers Selected ***"
      End If
   Else
      txtNme = "*** All Customers Selected ***"
   End If
   
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getreportcu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub

Private Sub txtSon_LostFocus()
   txtSon = Format(Abs(Val(txtSon)), SO_NUM_FORMAT)
   If Val(txtSon) > 0 Then
      cmbCst = "ALL"
      txtNme = "*** All Customers Selected ***"
      lblSoType = GetSalesOrderType()
   Else
      lblSoType = ""
   End If
   
End Sub



Private Function GetSalesOrderType()
   Dim RdoTyp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetSalesOrderType " & Val(txtSon) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTyp, ES_FORWARD)
   If bSqlRows Then
      GetSalesOrderType = "" & Trim(RdoTyp!SOTYPE)
   Else
      GetSalesOrderType = ""
   End If
   Set RdoTyp = Nothing
   Exit Function
   
DiaErr1:
   GetSalesOrderType = ""
   
End Function

Private Function GetAllocations() As Byte
   Dim RdoAlc As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT RAREF FROM RnalTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAlc, ES_FORWARD)
   If bSqlRows Then ClearResultSet RdoAlc
   GetAllocations = bSqlRows
   Set RdoAlc = Nothing
   
End Function
