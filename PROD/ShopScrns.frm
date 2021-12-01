VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ShopScrns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close Manufacturing Orders"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "R&uns"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Close Selected Runs"
      Top             =   3720
      Width           =   850
   End
   Begin VB.CommandButton cmdFil 
      Caption         =   "&Select"
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Fill The Grid"
      Top             =   3240
      Width           =   850
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtParts 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter Leading Characters Or Blank (Up To 50 Entries)"
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4395
      FormDesignWidth =   6855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Select Then Double Click Or Select And Press Enter To Add Or Remove Run"
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Site Is Currently Being Revised"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "ShopScrns.frx":0000
      ToolTipText     =   "Under Construction"
      Top             =   0
      Width           =   480
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Close These Runs"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   240
      Picture         =   "ShopScrns.frx":0442
      Top             =   3960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   600
      Picture         =   "ShopScrns.frx":07CC
      Top             =   3960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   11
      Tag             =   "Pa"
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblSelected 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed  From"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Tag             =   "Pa"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "By Date Range"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "ShopScrns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Dim sOrders(51, 4) As String
'0 = Selected "X" or ""
'1 = Compressed Mo
'2 = Str$(RUNNO)
'3 = Closed "X" or ""
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdClose_Click()
   Dim bByte As Byte
   Dim iList As Byte
   For iList = 1 To Grid1.Rows - 2
      If sOrders(iList, 0) = "X" Then bByte = 1
   Next
   If sOrders(iList, 0) = "X" Then bByte = 1
   
   If bByte = 0 Then
      MsgBox "At Least One MO To Close Must Be Selected.", _
         vbInformation, Caption
   Else
      'Test Closing
   End If
   'UnderCons.Show
   
End Sub

Private Sub cmdFil_Click()
   GetCompletedRuns
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then GetRunDates
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   FormatGrid
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopScrns = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtBeg.ToolTipText = "Earliest Found"
   txtEnd.ToolTipText = "Latest Found"
   
End Sub


Private Sub Grid1_Click()
   If Grid1.row > 1 Then Grid1.Col = 0
   
End Sub

Private Sub Grid1_DblClick()
   On Error Resume Next
   Grid1.Col = 0
   If Grid1.CellPicture = Chkyes.Picture Then
      Set Grid1.CellPicture = Chkno.Picture
      sOrders(Grid1.row, 0) = ""
   Else
      Set Grid1.CellPicture = Chkyes.Picture
      sOrders(Grid1.row, 0) = "X"
   End If
   
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   Grid1.Col = 1
   If Grid1.CellPicture = Chkyes.Picture Then
      Set Grid1.CellPicture = Chkno.Picture
      sOrders(Grid1.row, 0) = ""
   Else
      Set Grid1.CellPicture = Chkyes.Picture
      sOrders(Grid1.row, 0) = "X"
   End If
   
End Sub


Private Sub Image1_Click()
   'UnderCons.Show
   
End Sub

Private Sub Label1_Click()
   'UnderCons.Show
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub



Private Sub GetRunDates()
   Dim RdoDate As ADODB.Recordset
   sSql = "SELECT MIN(RUNCOMPLETE) AS BegDate FROM RunsTable " _
          & "WHERE RUNSTATUS='CO'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoDate!BegDate) Then txtBeg = _
                    Format(RdoDate!BegDate, "mm/dd/yy") Else _
                    txtBeg = Format(ES_SYSDATE, "mm/dd/yy")
      ClearResultSet RdoDate
   End If
   
   sSql = "SELECT MAX(RUNCOMPLETE) AS EndDate FROM RunsTable " _
          & "WHERE RUNSTATUS='CO'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoDate!EndDate) Then txtEnd = _
                    Format(RdoDate!EndDate, "mm/dd/yy") Else _
                    txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
      ClearResultSet RdoDate
   End If
   
   
End Sub

Public Sub GetCompletedRuns()
   Dim RdoRuns As ADODB.Recordset
   Dim bLen As Byte
   Dim iCounter As Integer
   Dim iRows As Integer
   Dim sPartRange As String
   sPartRange = Compress(txtParts)
   bLen = Len(sPartRange)
   If bLen = 0 Then bLen = 1
   FormatGrid
   iRows = 0
   Erase sOrders
   sSql = "SELECT RUNREF,RUNNO,RUNCOMPLETE,RUNYIELD,PARTREF,PARTNUM,PADESC FROM RunsTable,PartTable WHERE " _
          & "(RUNSTATUS='CO' AND LEFT(RUNREF," & bLen & ")>= '" & sPartRange & " ' " _
          & "AND RUNCOMPLETE BETWEEN '" & txtBeg & " 00:00' AND '" & txtEnd & " 00:00') " _
          & "AND PARTREF=RUNREF AND RUNYIELD>0 ORDER BY RUNCOMPLETE,RUNREF,RUNNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRuns, ES_FORWARD)
   If bSqlRows Then
      With RdoRuns
         Do Until .EOF
            iRows = iRows + 1
            iCounter = iCounter + 1
            If iRows > 1 Then Grid1.Rows = Grid1.Rows + 1
            sOrders(iRows, 0) = ""
            sOrders(iRows, 1) = "" & Trim(!RUNREF)
            sOrders(iRows, 2) = "" & Trim(str$(!Runno))
            sOrders(iRows, 3) = ""
            Grid1.row = iRows
            Grid1.Col = 0
            Set Grid1.CellPicture = Chkno.Picture
            Grid1.Col = 1
            Grid1.Text = "" & Format(!RUNCOMPLETE, "mm/dd/yy")
            Grid1.Col = 2
            Grid1.Text = "" & Trim(!PartNum)
            Grid1.Col = 3
            Grid1.Text = !Runno
            Grid1.Col = 4
            Grid1.Text = "" & Trim(!PADESC)
            Grid1.Col = 5
            Grid1.Text = "" & Format(!RUNYIELD, "######0")
            .MoveNext
         Loop
         ClearResultSet RdoRuns
         Grid1.Enabled = True
         lblSelected = iCounter
      End With
   Else
      Grid1.Enabled = False
   End If
   If iCounter > 0 Then cmdClose.Enabled = True Else _
                                           cmdClose.Enabled = False
   Set RdoRuns = Nothing
   
End Sub

Private Sub FormatGrid()
   Grid1.Clear
   With Grid1
      .Rows = 2
      .ColWidth(0) = 500
      .ColWidth(1) = 800
      .ColWidth(2) = 2000
      .ColWidth(3) = 500
      .ColWidth(4) = 1900
      .ColWidth(5) = 800
      .ColAlignment(2) = 0
      .ColAlignment(5) = 0
      .row = 0
      .Col = 0
      .Text = "Close"
      .Col = 1
      .Text = "Date"
      .Col = 2
      .Text = "Part Number"
      .Col = 3
      .Text = "Run"
      .Col = 4
      .Text = "Description"
      .Col = 5
      .Text = "Quantity"
      Grid1.row = 1
      Grid1.Col = 0
      Set Grid1.CellPicture = Chkno.Picture
   End With
   
End Sub
