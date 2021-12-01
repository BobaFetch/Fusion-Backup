VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form EstiESe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accept Estimates"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtBid 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.ComboBox txtAcp 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1250
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1320
      Width           =   1250
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "Fill The Grid"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Update Selections"
      Top             =   1320
      Width           =   875
   End
   Begin VB.ComboBox cmbCst 
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Bids"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   720
      Top             =   4680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4845
      FormDesignWidth =   7290
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Click Or Press Enter To Select - ESC To Cancel"
      Top             =   2040
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      GridLines       =   3
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
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "The List Will Contain Only Completed Estimates"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   300
      Width           =   3972
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   0
      Picture         =   "EstiESe04a.frx":07AE
      Top             =   4680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   240
      Picture         =   "EstiESe04a.frx":0B38
      Top             =   4680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimates From"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   2292
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rows Selected:"
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Accepted"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   2292
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   10
      Top             =   600
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Dates Ending"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   2292
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1092
   End
End
Attribute VB_Name = "EstiESe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/12/06 Corrected Grid
Option Explicit
Dim bOnLoad As Byte

Dim iTotalRows As Integer
Dim sBids(300, 3) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3504
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdSel_Click()
   FillGrid
   
End Sub




Private Sub cmdUpd_Click()
   Dim b As Byte
   Dim iList As Integer
   Dim sMsg As String
   Dim vDate As Variant
   
   b = 0
   For iList = 1 To iTotalRows
      If sBids(iList, 1) = "X" Then b = 1
   Next
   If b = 0 Then
      MsgBox "No Estimates Were Selected For Acceptance.", _
         vbInformation, Caption
   Else
      vDate = Format(txtAcp, "mm/dd/yyyy")
      sMsg = "This Procedure Marks All Selected Estimates " & vbCrLf _
             & "Completed (Where Not) And Accepted. Would    " & vbCrLf _
             & "You Like To Continue Accepting Estimates?"
      b = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If b = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         For iList = 1 To iTotalRows
            If sBids(iList, 1) = "X" Then
               If sBids(iList, 2) = "" Then
                  sSql = "UPDATE EstiTable SET BIDCOMPLETED='" & vDate & "'," _
                         & "BIDCOMPLETE=1,BIDACCEPT='" & vDate & "',BIDACCEPTED=1 " _
                         & "WHERE BIDREF=" & Val(sBids(iList, 0)) & " "
               Else
                  sSql = "UPDATE EstiTable SET BIDACCEPT='" & vDate & "'," _
                         & "BIDACCEPTED=1 WHERE BIDREF=" & Val(sBids(iList, 0)) & " "
               End If
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
            End If
         Next
         If clsADOCon.ADOErrNum = 0 Then
            MsgBox "Marked Estimates Completed And Accepted.", _
               vbInformation, Caption
            FillGrid
         Else
            MsgBox "Marked Estimates Were Not Completed Or Accepted.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
   
   
End Sub

Private Sub Form_Activate()
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
   With Grid1
      .Rows = 2
      .ColWidth(0) = 700
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 1200
      .ColWidth(4) = 2400
      .ColWidth(5) = 900
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .ColAlignment(4) = 0
      '.ColAlignment(5) = Use default 4/13/06
      .Row = 0
      .Col = 0
      .Text = "Accept"
      .Col = 1
      .Text = "Estimate"
      .Col = 2
      .Text = "Date "
      .Col = 3
      .Text = "Customer"
      .Col = 4
      .Text = "Part Number "
      .Col = 5
      .Text = "Quantity "
   End With
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
   Set EstiESe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtAcp = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BIDCUST,CUREF,CUNICKNAME FROM " _
          & "EstiTable,CustTable WHERE BIDCUST=CUREF ORDER BY BIDCUST"
   LoadComboBox cmbCst, 1
   cmbCst = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillGrid()
   Dim RdoRun As ADODB.Recordset
   Dim iList As Integer
   Dim sCust As String
   Dim vDate As Variant
   
   On Error Resume Next
   Grid1.Rows = 1
   Grid1.Row = 0
   
   On Error GoTo DiaErr1
   cmdUpd.Enabled = False
   Erase sBids
   If cmbCst <> "ALL" Then sCust = Compress(cmbCst) Else sCust = ""
   iTotalRows = 0
   vDate = Format(txtDte, "mm/dd/yyyy")
   sSql = "SELECT BIDREF,BIDNUM,BIDPRE,BIDPART,BIDCUST,BIDQUANTITY," _
          & "BIDDATE, BIDCOMPLETE,CUREF,CUNICKNAME,CUNAME,PARTREF,PARTNUM " _
          & "FROM EstiTable,CustTable,PartTable WHERE (BIDCUST=CUREF " _
          & "AND BIDPART=PARTREF AND BIDCANCELED=0 AND BIDACCEPTED=0 " _
          & "AND BIDCOMPLETE=1 AND BIDREF>=" & Val(txtBid) & " AND BIDCUST LIKE '" _
          & sCust & "%') AND BIDDATE<='" & vDate & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      With RdoRun
         Do Until .EOF
            iList = iList + 1
            If iList > 298 Then Exit Do
            sBids(iList, 0) = "" & Trim(!BIDNUM)
            sBids(iList, 1) = ""
            Grid1.Rows = iList + 1
            Grid1.Row = iList
            Grid1.Col = 0
            '                    If !BIDACCEPTED = 0 Then
            '                        sBids(iList, 2) = "X"
            '                        Set Grid1.CellPicture = Chkyes.Picture
            '                    Else
            sBids(iList, 2) = ""
            Set Grid1.CellPicture = Chkno.Picture
            '                    End If
            Grid1.Col = 1
            Grid1.Text = "" & Trim(!BIDPRE) & Trim(!BIDNUM)
            Grid1.Col = 2
            Grid1.Text = "" & Format(!BIDDATE, "mm/dd/yyyy")
            Grid1.Col = 3
            Grid1.Text = "" & Trim(!CUNICKNAME)
            Grid1.Col = 4
            Grid1.Text = "" & Trim(!PartNum)
            Grid1.Col = 5
            Grid1.Text = Format(!BidQuantity, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoRun
      End With
      If iList > 299 Then MsgBox "More Than 300 Rows Selected. " & vbCrLf _
         & "You May Wish To Narrow The Search.", _
         vbInformation, Caption
      iTotalRows = Grid1.Rows - 1
      lblRows = iTotalRows
      If iTotalRows > 0 Then cmdUpd.Enabled = True
      Grid1.Row = 1
      Grid1.Col = 0
      Grid1.SetFocus
   Else
      lblRows = 0
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

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Grid1.CellPicture = Chkyes.Picture Then
         Set Grid1.CellPicture = Chkno.Picture
         sBids(Grid1.Row, 1) = ""
      Else
         Set Grid1.CellPicture = Chkyes.Picture
         sBids(Grid1.Row, 1) = "X"
      End If
   End If
   
End Sub


Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grid1.Col = 0
   If Grid1.CellPicture = Chkyes.Picture Then
      Set Grid1.CellPicture = Chkno.Picture
      sBids(Grid1.Row, 1) = ""
   Else
      Set Grid1.CellPicture = Chkyes.Picture
      sBids(Grid1.Row, 1) = "X"
   End If
   
End Sub

Private Sub txtAcp_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtAcp_LostFocus()
   txtAcp = CheckDateEx(txtAcp)
   
End Sub


Private Sub txtBid_LostFocus()
   txtBid = Format(Abs(Val(txtBid)), "000000")
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   
End Sub



Private Sub GetOptions()
   On Error Resume Next
   txtBid = GetSetting("Esi2000", "EsiEngr", "EsAcp", txtBid)
   If Val(txtBid) = 0 Then GetFirstBid
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiEngr", "EsAcp", txtBid
   
End Sub

Private Sub GetFirstBid()
   Dim RdoFst As ADODB.Recordset
   sSql = "SELECT MIN(BIDREF) FROM EstiTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFst, ES_FORWARD)
   If bSqlRows Then
      With RdoFst
         If Not IsNull(.Fields(0)) Then
            txtBid = Format(.Fields(0), "000000")
         Else
            txtBid = "000001"
         End If
         ClearResultSet RdoFst
      End With
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getfirstbid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
