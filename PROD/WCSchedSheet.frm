VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form WCSchedSheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALL Work Center MO Schedules"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   16410
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "WCSchedSheet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optData 
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optReport 
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optFrom 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSel 
      Cancel          =   -1  'True
      Caption         =   "S&elect"
      Height          =   315
      Left            =   6120
      TabIndex        =   2
      ToolTipText     =   "Fill The Grid"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   360
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8775
      Left            =   220
      TabIndex        =   3
      ToolTipText     =   "Double Click Or Press Enter To Edit Entry"
      Top             =   840
      Width           =   16185
      _ExtentX        =   28549
      _ExtentY        =   15478
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   100
      WordWrap        =   -1  'True
      Enabled         =   0   'False
      AllowUserResizing=   2
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   15240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   14880
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9990
      FormDesignWidth =   16410
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2520
      TabIndex        =   12
      Top             =   9600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Sched Complete from"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "AWI Special"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "WCSchedSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/8/05 Removed Jet and created report tables
'10/20/05 Re-added the operation and changed the resizing of the form
Option Explicit

Public bOnLoad As Byte
Dim bRefreshed As Byte
Dim lSonumber As Long
Dim iItem As Integer
Dim sRev As String
Dim sDesc As String

Dim sPkPart As String
Dim sPkComt As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdSel_Click()
   FillGrid
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      
      Dim sql As String
      
      Grid1.Clear
      Grid1.Rows = 2
      
'      sql = "SELECT DISTINCT a.OPCENTER OpCenter FROM RnopTable a, " _
'         & " (SELECT top(10) OPCENTER, MIN(RUNSCHED) RUNSCHED FROM RnopTable, RunsTable " _
'         & " WHERE (OPREF = RUNREF And OPRUN = Runno)" _
'            & " AND OPCOMPLETE=0 GROUP BY OPCENTER)" _
'            & " as f WHERE f.OPCENTER =a.OPCENTER"
   
      FillGridHeader
      
      bOnLoad = 0
   End If
   If optFrom.Value = vbChecked Then
      optFrom.Value = vbUnchecked
      Unload WCSchedSheet1a
   End If
   MouseCursor 0
   
End Sub

Private Function FillGridHeader()
   Dim RdoCnt As ADODB.Recordset
   Dim ColCnt As Integer
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT WorkCenter FROM WCSchView"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
   If bSqlRows Then
      ColCnt = 0
      With RdoCnt
         Do Until .EOF
            
            Grid1.row = 0
            Grid1.Col = ColCnt
            Grid1.Text = !WorkCenter
            
            .MoveNext
            ColCnt = ColCnt + 1
         Loop
         ClearResultSet RdoCnt
      End With
   End If
   
   Set RdoCnt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "FillGridHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   

End Function

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   
   With Grid1
      .Rows = 2
      .ColWidth(0) = 1550
      .ColWidth(1) = 1550
      .ColWidth(2) = 1550
      .ColWidth(3) = 1550
      .ColWidth(4) = 1550
      .ColWidth(5) = 1550
      .ColWidth(6) = 1550
      .ColWidth(7) = 1550
      .ColWidth(8) = 1550
      .ColWidth(9) = 1550
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .ColAlignment(4) = 0
      .ColAlignment(6) = 0
      .ColAlignment(7) = 0
      .ColAlignment(9) = 0
      
      End With
   bRefreshed = 0
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   Grid1.ToolTipText = "Contains Up To 300 Entries. " & Grid1.ToolTipText
   'Janet 12/2/04
   txtBeg = Format(ES_SYSDATE, "mm") & "/01/" & Format(ES_SYSDATE, "yy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   
End Sub

Private Sub Grid1_DblClick()
   On Error GoTo DiaErr1
      
   Dim strVal As String
   Dim strWC As String
   Dim strSetup As String
   Dim strCmt As String
   
   Dim arVal As Variant
   
   strVal = Grid1.Text
   arVal = Split(strVal, "  ")
   
   Grid1.row = 0
   strWC = Grid1.Text
   
   Dim RdoSet As ADODB.Recordset
   
   sSql = "SELECT OPSUHRS, OPCOMT FROM RnopTable  " _
            & " WHERE OPREF = '" & Compress(arVal(0)) & "' AND " _
            & " OPRUN = '" & arVal(1) & "' " _
            & " AND OPNO = '" & arVal(2) & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_STATIC)
   
   If bSqlRows Then
      With RdoSet
         On Error Resume Next
         strSetup = !OPSUHRS
         If (strSetup = "") Then strSetup = "0.0000"
         
         strCmt = Trim(!OPCOMT)
      End With
   End If
   Set RdoSet = Nothing
   
   
   WCSchedSheet1a.cmbWcn = strWC
   WCSchedSheet1a.txtPri = arVal(5)
   WCSchedSheet1a.lblMon = arVal(0)
   WCSchedSheet1a.lblRun = arVal(1)
   WCSchedSheet1a.lblOpno = arVal(2)
   'WCSchedSheet1a.lblQty = arVal(4)
   
   WCSchedSheet1a.txtSetup = strSetup
   WCSchedSheet1a.txtCmt = strCmt
   WCSchedSheet1a.lblQty = arVal(4)
   
   WCSchedSheet1a.Show
   WCSchedSheet1a.lblWcn = strWC
   
   Exit Sub
   
DiaErr1:
   sProcName = "GridDblClick"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      On Error Resume Next
      Dim strVal As String
      Dim strWC As String
      
      
      Dim arVal As Variant
      
      strVal = Grid1.Text
      arVal = Split(strVal, "  ")
      
      Grid1.row = 0
      strWC = Grid1.Text
      
      WCSchedSheet1a.cmbWcn = strWC
      'WCSchedSheet1a.txtPri = arVal(5)
      WCSchedSheet1a.lblMon = arVal(0)
      WCSchedSheet1a.lblRun = arVal(1)
      WCSchedSheet1a.lblOpno = arVal(2)
      WCSchedSheet1a.lblQty = arVal(4)
      
      WCSchedSheet1a.Show
      WCSchedSheet1a.lblWcn = strWC
   End If
   
End Sub


Private Sub optData_Click()
   'if the data in the PopUp has been changed
   
End Sub

Private Sub optFrom_Click()
   'Shows edit
   'Refill when closing the 2nd form
   If optFrom.Value = vbUnchecked And optData.Value = vbChecked Then
      optData.Value = vbUnchecked
      FillGrid
   End If
   
End Sub


Private Sub txtBeg_DropDown()
    ShowCalendar Me
End Sub

Private Sub txtBeg_LostFocus()
    txtEnd = CheckDate(txtEnd)
End Sub



Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
   
End Sub



Private Sub FillGrid()
   Dim RdoGrd As ADODB.Recordset
   Dim b As Byte
   Dim iList As Integer
   Dim A As Integer
   Dim C As Integer
   Dim sCenter As String
   Dim sNick As String
   Dim sSales As String
   Dim sSoPon As String
   Dim lblTotal As Long
   
   
   'FillGridHeader sql
   Grid1.Clear
   Grid1.Rows = 1
   Grid1.row = 0
   
   On Error Resume Next
   A = 10
   
   'sSql = "RptROLT '" & sPartNum & "','" & sInitials & "'"
   'RdoCon.Execute sSql, rdExecDirect


   sSql = "PivotWCSchDate '" & Format(txtBeg, "mm/dd/yy") & "','" & Format(txtEnd, "mm/dd/yy") & "'"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      
      FillGridHdRec RdoGrd
      
      lblInfo = "Filling Grid."
      lblInfo.Visible = True
      lblInfo.Refresh
      prg1.Value = A
      prg1.Visible = True
      
      With RdoGrd
         Do Until .EOF
            iList = iList + 1
            C = C + 1
            If C = 10 Then
               A = A + 7
               C = 0
               If A > 95 Then A = 95
               prg1.Value = A
            End If
            
            Grid1.Rows = iList + 1
            
            Grid1.Col = 0
            Grid1.row = iList
            Grid1.RowHeight(iList) = 1000
            Grid1.Text = Trim(RdoGrd.Fields(1))
            
            Grid1.Col = 1
            Grid1.Text = Trim(RdoGrd.Fields(2))
            
            Grid1.Col = 2
            Grid1.Text = Trim(RdoGrd.Fields(3))
            
            Grid1.Col = 3
            Grid1.Text = Trim(RdoGrd.Fields(4))
            
            Grid1.Col = 4
            Grid1.Text = Trim(RdoGrd.Fields(5))
            
            Grid1.Col = 5
            Grid1.Text = Trim(RdoGrd.Fields(6))
            
            Grid1.Col = 6
            Grid1.Text = Trim(RdoGrd.Fields(7))
            
            Grid1.Col = 7
            Grid1.Text = Trim(RdoGrd.Fields(8))
            
            Grid1.Col = 8
            Grid1.Text = Trim(RdoGrd.Fields(9))
            
            Grid1.Col = 9
            Grid1.Text = Trim(RdoGrd.Fields(10))
            
            
            lblTotal = iList
            'lblTotal.Refresh
            .MoveNext
         Loop
         ClearResultSet RdoGrd
      End With
      If Grid1.Rows > 1 Then Grid1.Enabled = True _
                                             Else Grid1.Enabled = False
      prg1.Value = 100
      Grid1.Col = 0
      Grid1.row = 1
      On Error Resume Next
      Grid1.SetFocus
   Else
      MouseCursor 0
      lblTotal = 0
      MsgBox "No Open Operations Found.", _
         vbInformation, Caption
   End If
   prg1.Visible = False
   lblInfo.Visible = False
   MouseCursor 0
   Set RdoGrd = Nothing
   
End Sub

Private Function FillGridHdRec(RdoGrd As ADODB.Recordset)
      
      Grid1.row = 0
      Grid1.Col = 0
      Grid1.Text = RdoGrd.Fields(1).Name
      Grid1.Col = 1
      Grid1.Text = RdoGrd.Fields(2).Name
      Grid1.Col = 2
      Grid1.Text = RdoGrd.Fields(3).Name
      Grid1.Col = 3
      Grid1.Text = RdoGrd.Fields(4).Name
      Grid1.Col = 4
      Grid1.Text = RdoGrd.Fields(5).Name
      Grid1.Col = 5
      Grid1.Text = RdoGrd.Fields(6).Name
      Grid1.Col = 6
      Grid1.Text = RdoGrd.Fields(7).Name
      Grid1.Col = 7
      Grid1.Text = RdoGrd.Fields(8).Name
      Grid1.Col = 8
      Grid1.Text = RdoGrd.Fields(9).Name
      Grid1.Col = 9
      Grid1.Text = RdoGrd.Fields(10).Name

End Function

