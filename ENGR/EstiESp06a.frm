VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requests For Quotation (Report)"
   ClientHeight    =   4395
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cbRFQNoEst 
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   3840
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   2520
      Width           =   4095
      Begin VB.OptionButton optAcc 
         Caption         =   "Complete"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   200
         Width           =   1095
      End
      Begin VB.OptionButton optAcc 
         Caption         =   "Incomplete"
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   21
         Top             =   200
         Width           =   1335
      End
      Begin VB.OptionButton optAcc 
         Caption         =   "ALL"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   200
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   18
      Top             =   2160
      Width           =   1250
   End
   Begin VB.ComboBox cmbBeg 
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   2160
      Width           =   1250
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "EstiESp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Picture         =   "EstiESp06a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox OptExt 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   3240
      Value           =   1  'Checked
      Width           =   525
   End
   Begin VB.ComboBox cmbCst 
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select A Customer"
      Top             =   960
      Width           =   1555
   End
   Begin VB.ComboBox cmbRfq 
      Enabled         =   0   'False
      Height          =   288
      ItemData        =   "EstiESp06a.frx":0938
      Left            =   2160
      List            =   "EstiESp06a.frx":093A
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter RFQ Number"
      Top             =   1680
      Width           =   2040
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "EstiESp06a.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "EstiESp06a.frx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   3480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4395
      FormDesignWidth =   7110
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   29
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Through"
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "RFQ's Due From"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Corresponding Estimate"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Include RFQ's with no"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Show RFQ's"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   252
      Index           =   5
      Left            =   4800
      TabIndex        =   14
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   1680
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   252
      Index           =   4
      Left            =   4800
      TabIndex        =   13
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimates"
      Height          =   165
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   3240
      Width           =   1665
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   10
      Top             =   1320
      Width           =   3252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   960
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ Number"
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   1680
      Width           =   1332
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   675
      TabIndex        =   7
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "EstiESp06a"
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

Private Sub cmbBeg_DropDown()
    ShowCalendarEx Me
End Sub

Private Sub cmbBeg_LostFocus()
    If Len(Trim(cmbBeg)) = 0 Then cmbBeg = "ALL"
    If cmbBeg <> "ALL" Then cmbBeg = CheckDateEx(cmbBeg)
End Sub

Private Sub cmbCst_Click()
   If cmbCst = "" Or cmbCst = "ALL" Then
      txtNme = "All Customers Selected"
      cmbCst = "ALL"
      cmbRfq = "ALL"
      cmbRfq.Enabled = False
   Else
      FindCustomer Me, cmbCst
      FillCustomerRFQs Me, cmbCst, False
      If cmbRfq.ListCount = 0 Then cmbRfq = "ALL"
      cmbRfq.Enabled = True
   End If
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If cmbCst = "" Or cmbCst = "ALL" Then
      txtNme = "All Customers Selected"
      cmbCst = "ALL"
      cmbRfq = "ALL"
      cmbRfq.Enabled = False
   Else
      FindCustomer Me, cmbCst
      FillCustomerRFQs Me, cmbCst, False
      If cmbRfq.ListCount = 0 Then cmbRfq = "ALL"
      cmbRfq.Enabled = True
   End If
   
End Sub


Private Sub cmbEnd_DropDown()
    ShowCalendarEx Me
End Sub

Private Sub cmbEnd_LostFocus()
    If Len(Trim(cmbEnd)) = 0 Then cmbEnd = "ALL"
    If cmbEnd <> "ALL" Then cmbEnd = CheckDateEx(cmbEnd)
End Sub

Private Sub cmbRfq_LostFocus()
   cmbRfq = CheckLen(cmbRfq, 14)
   If cmbRfq = "" Then cmbRfq = "ALL"
   
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

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      cmbCst.AddItem "ALL"
      FillCustomers
      cmbCst = "ALL"
      cmbRfq = "ALL"
      txtNme = "All Customers Selected"
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
   Set EstiESp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sRfq As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sBDate, sEDate As String

   If Trim(cmbCst) <> "ALL" Then sCust = Compress(cmbCst)
   If Trim(cmbRfq) <> "ALL" Then sRfq = cmbRfq
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   If Not IsDate(cmbBeg) Then sBDate = "1995,01,01" Else sBDate = Format(cmbBeg, "yyyy,mm,dd")
   If Not IsDate(cmbEnd) Then sEDate = "2024,12,31" Else sEDate = Format(cmbEnd, "yyyy,mm,dd")
   
   sCustomReport = GetCustomReport("enges06")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowExDesc"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer " & CStr(cmbCst _
                        & " And RFQ " & cmbRfq) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add OptExt.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'sSql = "{EstiTable.BIDCUST} LIKE '" & sCust & "*' " _
   '       & "AND {EstiTable.BIDRFQ} LIKE '" & sRfq & "*' "
   
   
   
   sSql = "({RfqsTable.RFQCUST} LIKE '" & Trim(sCust) & "*' " _
          & "AND {RfqsTable.RFQREF} LIKE '" & Trim(sRfq) & "*' )"
          
   sSql = sSql & " AND ( "
   sSql = sSql & " ( {RfqsTable.RFQDUE} IN Date(" & sBDate & ") To Date(" & sEDate & ") "
   If optAcc(1).value = True Then sSql = sSql & " AND {RfqsTable.RFQCOMPLETE}=1 )" _
      Else If optAcc(2).value = True Then sSql = sSql & " AND {RfqsTable.RFQCOMPLETE} = 0 )" Else sSql = sSql & ") "
   
   'If cbRFQNoEst.value = 1 Then sSql = sSql & " OR ( IsNull({EstiTable.BIDRFQ}) ) "
      
   sSql = sSql & ")"
          
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
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim I As Integer
   
   SaveSetting "Esi2000", "EsiProd", "Pes06", lblPrinter
   
   sOptions = Left(cmbCst & Space(10), 10)
   sOptions = sOptions & Left(cmbRfq & Space(20), 20)
   sOptions = sOptions & Left(cmbBeg & Space(8), 10)
   sOptions = sOptions & Left(cmbEnd & Space(8), 10)
   For I = optAcc.LBound To optAcc.UBound
        If optAcc(I).value = True Then sOptions = sOptions & Right("0" & Trim(str(I)), 1)
   Next I
   sOptions = sOptions & Right("0" & Trim(str(OptExt.value)), 1)
   sOptions = sOptions & Right("0" & Trim(str(cbRFQNoEst.value)), 1)
   
   SaveSetting "Esi2000", "EsiEngr", "EstiESp06a", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pes06", lblPrinter)
   
    sOptions = GetSetting("Esi2000", "EsiEngr", "EstiESp06a", "ALL                 ALL       ALL     ALL     011")
        
    On Error Resume Next
    cmbCst = Trim(Left(sOptions, 10))
    cmbRfq = Trim(Mid(sOptions, 11, 20))
    cmbBeg = Trim(Mid(sOptions, 31, 10))
    cmbEnd = Trim(Mid(sOptions, 41, 10))
    optAcc(Val(Mid(sOptions, 51, 1))) = vbChecked
    OptExt.value = Val(Mid(sOptions, 52, 1))
    cbRFQNoEst = Val(Mid(sOptions, 53, 1))
   
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

