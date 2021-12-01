VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ShopSHe06a 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Release Manufacturing Orders"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "&None"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Unmark All (None)"
      Top             =   1080
      Width           =   875
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&ALL"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Mark All"
      Top             =   1080
      Width           =   875
   End
   Begin VB.OptionButton optSrt 
      Caption         =   "Date"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton optSrt 
      Caption         =   "MO"
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Update And Save Selections"
      Top             =   960
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "Fill The Grid"
      Top             =   560
      Width           =   875
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4935
      FormDesignWidth =   6975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3252
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Click Or Use SpaceBar To Select - ESC To Cancel"
      Top             =   1440
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   240
      Picture         =   "ShopSHe06a.frx":07AE
      Top             =   4800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   0
      Picture         =   "ShopSHe06a.frx":0B38
      Top             =   4800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(On Or Before)"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO's Scheduled To Start "
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "ShopSHe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/5/04 Static Document List reminder
'4/28/06 Moved the Check to the right and re-aligned the form.
'6/19/06 Fixed errant Grid CheckBox
Option Explicit
Dim bDocList As Byte
Dim iTotalRows As Integer
Dim vGridRows(300, 3) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdAll_Click()
   Dim iRow As Integer
   Grid1.Col = 0
   For iRow = 1 To Grid1.Rows - 1
      Grid1.row = iRow
      Set Grid1.CellPicture = Chkyes.Picture
      vGridRows(Grid1.row, 2) = "X"
   Next
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4106
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdNone_Click()
   Dim iRow As Integer
   Grid1.Col = 0
   For iRow = 1 To Grid1.Rows - 1
      Grid1.row = iRow
      Set Grid1.CellPicture = Chkno.Picture
      vGridRows(Grid1.row, 2) = " "
   Next
   
End Sub

Private Sub cmdSel_Click()
   FillGrid
   
End Sub

Private Sub cmdUpd_Click()
   UpdateOrders
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   With Grid1
      .Rows = 2
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 1
      .ColAlignment(4) = 0
      .ColWidth(0) = 500
      .ColWidth(1) = 2950
      .ColWidth(2) = 800
      .ColWidth(3) = 1200
      .ColWidth(4) = 1400
      .row = 0
      .Col = 0
      .Text = "RL"
      .Col = 1
      .Text = "Manufacturing Order"
      .Col = 2
      .Text = "Run"
      .Col = 3
      .Text = "Sched Start"
      .Col = 4
      .Text = "Doc List Assigned"
   End With
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   'txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtEnd = ""
End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      Grid1.Col = 0
      If Grid1.CellPicture = Chkyes.Picture Then
         Set Grid1.CellPicture = Chkno.Picture
         vGridRows(Grid1.row, 2) = " "
      Else
         Set Grid1.CellPicture = Chkyes.Picture
         vGridRows(Grid1.row, 2) = "X"
      End If
      Grid1.Col = 1
   End If
   
End Sub


Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Grid1.Col = 0
   If Grid1.CellPicture = Chkyes.Picture Then
      Set Grid1.CellPicture = Chkno.Picture
      vGridRows(Grid1.row, 2) = " "
   Else
      Set Grid1.CellPicture = Chkyes.Picture
      vGridRows(Grid1.row, 2) = "X"
   End If
   Grid1.Col = 1
   
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckDateEx(txtEnd)
   
End Sub



Private Sub FillGrid()
   Dim RdoRun As ADODB.Recordset
   Dim iList As Integer
   Dim vDate As Variant
   
   On Error Resume Next
   Grid1.Rows = 1
   iList = 0
   bDocList = 0
   On Error GoTo DiaErr1
   cmdUpd.Enabled = False
   iTotalRows = 0
   vDate = Format(txtEnd, "mm/dd/yyyy")
   Erase vGridRows
   cmdAll.Enabled = False
   cmdNone.Enabled = False
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNSTART," _
          & "PARTREF,PARTNUM,PADOCLISTREF FROM RunsTable,PartTable " _
          & "WHERE (RUNREF=PARTREF AND RUNSTART" _
          & "<='" & vDate & "' AND RUNSTATUS='SC') "
   If optSrt(1).Value = True Then
      sSql = sSql & "ORDER BY RUNSTART,RUNREF,RUNNO "
   Else
      sSql = sSql & "ORDER BY RUNREF,RUNSTART,RUNNO "
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         Do Until .EOF
            iList = iList + 1
            If iList > 200 Then Exit Do
            Grid1.Rows = iList + 1
            Grid1.Col = 0
            Set Grid1.CellPicture = Chkno.Picture
            Grid1.Col = 1
            Grid1.row = iList
            vGridRows(iList, 0) = "" & Trim(!RUNREF)
            vGridRows(iList, 1) = !Runno
            vGridRows(iList, 2) = ""
            Grid1.Text = "" & Trim(!PartNum)
            Grid1.Col = 2
            Grid1.Text = "" & Trim(!Runno)
            Grid1.Col = 3
            Grid1.Text = Format(!RUNSTART, "mm/dd/yy")
            Grid1.Col = 4
            If Trim(!PADOCLISTREF) <> "" Then
               Grid1.Text = "No"
            Else
               bDocList = 1
               Grid1.Text = "Yes"
            End If
            .MoveNext
         Loop
         ClearResultSet RdoRun
      End With
      cmdAll.Enabled = True
      cmdNone.Enabled = True
      iTotalRows = Grid1.Rows
      Grid1.Col = 0
      For iList = 1 To Grid1.Rows - 1
         Grid1.row = iList
         Set Grid1.CellPicture = Chkno.Picture
      Next
      If iTotalRows > 0 Then cmdUpd.Enabled = True
      Grid1.row = 1
      Grid1.Col = 0
      Grid1.SetFocus
   End If
   MouseCursor 0
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateOrders()
   Dim bResponse As Byte
   Dim iList As Integer
   Dim sMsg As String
   
   sMsg = "Release All Marked Manufacturing Orders?"
   If bDocList = 1 Then
      sMsg = sMsg & vbCr & "Note: You Should Check And Update  " & vbCr _
             & "Document Lists Where Necessary."
   End If
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmdUpd.Enabled = False
      MouseCursor 13
      For iList = 1 To iTotalRows
         If vGridRows(iList, 2) = "X" Then
            sSql = "UPDATE RunsTable SET RUNSTATUS='RL' " _
                   & "WHERE RUNREF='" & vGridRows(iList, 0) & "' " _
                   & "AND RUNNO=" & vGridRows(iList, 1) & " "
            clsADOCon.ExecuteSQL sSql
         End If
      Next
      MouseCursor 0
      MsgBox "All Marked MO's Have Been Released.", _
         vbInformation, Caption
      FillGrid
   Else
      CancelTrans
   End If
   
End Sub

Private Sub SaveOptions()
   On Error Resume Next
   SaveSetting "Esi2000", "EsiProd", "smorl", optSrt(1).Value
   
End Sub

Private Sub GetOptions()
   On Error Resume Next
   optSrt(1).Value = GetSetting("Esi2000", "EsiProd", "smorl", optSrt(1).Value)
   
End Sub
