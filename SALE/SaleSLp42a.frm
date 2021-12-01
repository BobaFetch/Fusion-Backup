VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp42a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Acknowledgements"
   ClientHeight    =   3495
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp42a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCan 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   2520
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select or Enter Sales Order Number (List Contains Last 3 Years Up To 500 Enties)"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   18
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "SaleSLp42a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SaleSLp42a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CheckBox optRem 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtEnd 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Tag             =   "3"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7200
      Top             =   3960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3495
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Canceled Items"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   3120
      Width           =   1668
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2280
      TabIndex        =   20
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2280
      TabIndex        =   19
      Top             =   1760
      Width           =   3852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Printed Only) - Disabled"
      Height          =   288
      Index           =   10
      Left            =   3360
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   3348
   End
   Begin VB.Label lblEnd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2280
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lblPre 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   288
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Sales Order"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1785
   End
End
Attribute VB_Name = "SaleSLp42a"
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
Dim bGoodBeg As Byte
Dim bGoodEnd As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   'Dim sYear As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbSon.Clear
   'iList = Format(Now, "yyyy")
   'iList = iList - 3
   'sYear = Trim$(iList) & "-" & Format(Now, "mm-dd")
   'sSql = "Qry_FillSalesOrders '" & sYear & "'"
   sSql = "Qry_FillSalesOrders '" & DateAdd("yyyy", -3, Now) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      iList = -1
      With RdoCmb
         lblPre = "" & Trim(!SOTYPE)
         cmbSon = Format(!SoNumber, SO_NUM_FORMAT)
         Do Until .EOF
            iList = iList + 1
            If iList > 999 Then Exit Do
            AddComboStr cmbSon.hWnd, Format$(!SoNumber, SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      MouseCursor 0
      MsgBox "No Sales Orders Where Found.", vbInformation, Caption
      Exit Sub
   End If
   Set RdoCmb = Nothing
   MouseCursor 0
   If cmbSon.ListCount > 0 Then bGoodBeg = GetSalesOrder()
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

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "sl42", sOptions)
   If Len(sOptions) > 0 Then
      optExt.Value = Val(Left(sOptions, 1))
      optCmt.Value = Val(Mid(sOptions, 2, 1))
      optRem.Value = Val(Mid(sOptions, 3, 1))
   End If
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optCmt.Value) _
              & RTrim(optRem.Value)
   SaveSetting "Esi2000", "EsiSale", "sl42", Trim(sOptions)
   
End Sub

Private Sub cmbSon_Click()
   bGoodBeg = GetSalesOrder()
   
End Sub

Private Sub cmbSon_LostFocus()
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   bGoodBeg = GetSalesOrder()
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2151
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then FillCombo
   bOnLoad = 0
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
   Set SaleSLp42a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim lSoNumber As Long
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo Sle42
   lSoNumber = Val(cmbSon)
   aFormulaName.Add "ShowComments"
   aFormulaName.Add "ShowExDescription"
   aFormulaName.Add "ShowRemarks"
   aFormulaValue.Add optCmt.Value
   aFormulaValue.Add optExt.Value
   aFormulaValue.Add optRem.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleco42")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{SohdTable.SONUMBER}=" & str$(lSoNumber) & " "
   If optCan.Value = vbUnchecked Then sSql = sSql _
                     & "AND {SoitTable.ITCANCELED}=0"
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
   
Sle42:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Sle42b
Sle42b:
   DoModuleErrors Me
   
End Sub

Private Sub lblNme_Change()
   If Trim(lblCst) = "" Then
      lblNme = "*** The Sales Order Was Not Found ***"
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optCan_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optCmt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   If bGoodBeg Then
      PrintReport
   Else
      MsgBox "Sales Order Wasn't Found.", vbInformation, Caption
   End If
   
End Sub

Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   If bGoodBeg Then
      PrintReport
   Else
      MsgBox "Sales Order Wasn't Found.", vbInformation, Caption
   End If
   
End Sub

Private Sub optRem_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optRem_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckLen(txtEnd, SO_NUM_SIZE)
   txtEnd = Format(Abs(Val(txtEnd)), SO_NUM_FORMAT)
   
End Sub

Private Function GetSalesOrder() As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   GetSalesOrder = 0
   sSql = "Qry_GetSalesOrderCustomer " & Val(cmbSon)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         lblPre = Trim(!SOTYPE)
         lblCst = "" & Trim(!CUNICKNAME)
         lblNme = "" & Trim(!CUNAME)
         GetSalesOrder = 1
      End With
      ClearResultSet RdoGet
   Else
      lblPre = ""
      lblCst = ""
      lblNme = ""
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   GetSalesOrder = False
   
End Function
