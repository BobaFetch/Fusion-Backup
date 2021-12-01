VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form RoutRTf08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Routing Operation Comments"
   ClientHeight    =   8625
   ClientLeft      =   3015
   ClientTop       =   1560
   ClientWidth     =   9360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8625
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCopy 
      Cancel          =   -1  'True
      Caption         =   "Copy"
      Height          =   435
      Left            =   7920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1230
   End
   Begin VB.TextBox txtCmt 
      Height          =   3585
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   "9"
      Text            =   "RoutRTf08a.frx":0000
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   4920
      Width           =   9015
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTf08a.frx":0007
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   1170
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   8160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2760
      Top             =   8520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8625
      FormDesignWidth =   9360
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Click To Select Or Scroll And Press Enter"
      Top             =   1560
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   375
      ScrollBars      =   2
      AllowUserResizing=   1
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
      Caption         =   "Comments"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   630
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label txtDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1170
      TabIndex        =   1
      Top             =   960
      Width           =   3075
   End
End
Attribute VB_Name = "RoutRTf08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/29/05 Fixed null value in GetOperations
Option Explicit

Dim AdoCmdStm As ADODB.Command
Dim AdoCmdStmOp As ADODB.Command

'Dim RdoStm As rdoQuery
'Dim RdoStmOp As rdoQuery

Dim bGoodRout As Byte
Dim bOnLoad As Byte

Dim iTotalOps As Integer
Dim iIndex As Integer

Dim strOpNum As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbRte_Click()
   bGoodRout = GetRout(False)
   ' Clear the grid
   ClearGrid
   GetOperations (False)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCan_Click
   
End Sub



Private Sub cmdCopy_Click()
   Clipboard.SetText txtCmt
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillRoutings
      If Len(sCurrRout) Then cmbRte = sCurrRout
      bGoodRout = GetRout(False)
      GetOperations (False)
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   On Error Resume Next
   sSql = "SELECT OPREF,OPNO,OPSHOP, OPCENTER, OPCOMT FROM RtopTable WHERE OPREF= ?"
   'Set RdoStm = RdoCon.CreateQuery("", sSql)
   
   Set AdoCmdStm = New ADODB.Command
   AdoCmdStm.CommandText = sSql
   
   Dim pPrtRef As ADODB.Parameter
   Set pPrtRef = New ADODB.Parameter
   pPrtRef.Type = adChar
   pPrtRef.Size = 30
   AdoCmdStm.Parameters.Append pPrtRef
   
   
   sSql = "SELECT OPCOMT FROM RtopTable WHERE OPREF= ? AND OPNO= ?"
   'Set RdoStmOp = RdoCon.CreateQuery("", sSql)
   
   Set AdoCmdStmOp = New ADODB.Command
   AdoCmdStmOp.CommandText = sSql
   
   Dim pRnPrt As ADODB.Parameter
   Set pRnPrt = New ADODB.Parameter
   pRnPrt.Type = adChar
   pRnPrt.Size = 30
   AdoCmdStmOp.Parameters.Append pRnPrt
   
   Dim pRunNO As ADODB.Parameter
   Set pRunNO = New ADODB.Parameter
   pRunNO.Type = adInteger
   AdoCmdStmOp.Parameters.Append pRunNO
   
   
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .Row = 0
      .Col = 0
      .Text = "Op No"
      .ColWidth(0) = 750
      .Col = 1
      .Text = "Shop"
      .ColWidth(1) = 1500
      .Col = 2
      .Text = "Work Center"
      .ColWidth(2) = 1500
      .Col = 3
      .Text = "Comment"
      .ColWidth(3) = 6000
      .Col = 0
   End With
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdoCmdStm = Nothing
   Set AdoCmdStmOp = Nothing
   
   FormUnload
   Set RoutRTf08a = Nothing
   
End Sub



Private Function GetRout(bGetOps As Byte) As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   sRout = Compress(cmbRte)
   GetRout = False
   
   On Error GoTo DiaErr1
   sSql = "Qry_GetToolRout '" & sRout & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte)
   If bSqlRows Then
      With RdoRte
         GetRout = True
         cmbRte = "" & Trim(!RTNUM)
         txtDsc = "" & Trim(!RTDESC)
         sCurrRout = sRout
         ClearResultSet RdoRte
      End With
      If bGetOps Then GetOperations (False)
   Else
      cmbRte = ""
      txtDsc = ""
      sCurrRout = ""
      GetRout = False
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub grd_Click()
   Grd.Col = 0
   'iIndex = Grd.Row
   strOpNum = Grd.Text
   GetThisOp
   
End Sub

Private Function GetOperations(bOpen As Byte)
   Dim RdoOps As ADODB.Recordset
   Dim iList As Integer
   Dim iRows As Integer
   Dim sString As String
   Dim sComment As String
   
   iTotalOps = 0
   iIndex = 0
   iList = -1
   iRows = 0
   On Error GoTo DiaErr1
   'RdoStm(0) = Compress(cmbRte)
   AdoCmdStm.Parameters(0) = Compress(cmbRte)
   bSqlRows = clsADOCon.GetQuerySet(RdoOps, AdoCmdStm, ES_STATIC)
   If bSqlRows Then
      With RdoOps
         Do Until .EOF
            iList = iList + 1
            iTotalOps = iTotalOps + 1
            iRows = iRows + 1
            If iRows > 1 Then Grd.Rows = Grd.Rows + 1
            Grd.Row = iRows
            Grd.Col = 0
            Grd.Text = Format(!OPNO, "000")
            Grd.Col = 1
            Grd.Text = "" & Trim(!OPSHOP)
            Grd.Col = 2
            
            Grd.Text = "" & Trim(!OPCENTER)
            Grd.Col = 3
            
            sString = "" & Trim(!OPCOMT)
            sComment = sString
            sString = Replace(sString, vbCrLf, " ")
            Grd.Text = sString
            
            If (iRows = 1) Then txtCmt = sComment
            
            .MoveNext
         Loop
         ClearResultSet RdoOps
      End With
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getopera"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetThisOp()
   Dim RdoRes2 As ADODB.Recordset
   Dim A As Integer
   On Error Resume Next
   
   
   'RdoStmOp(0) = Compress(cmbRte)
   'RdoStmOp(1) = Val(strOpNum)
   AdoCmdStmOp.Parameters(0) = Compress(cmbRte)
   AdoCmdStmOp.Parameters(1) = Val(strOpNum)
   
   bSqlRows = clsADOCon.GetQuerySet(RdoRes2, AdoCmdStmOp, ES_STATIC)
   If bSqlRows Then
      With RdoRes2
         txtCmt = "" & Trim(!OPCOMT)
         
         If Right(txtCmt, 1) = vbLf And Right(txtCmt, 1) = vbCr Then
            If Len(txtCmt) > 1 Then
               A = Len(txtCmt)
               txtCmt = (Left$(txtCmt, A - 1))
            Else
               txtCmt = ""
            End If
         End If
      End With
   End If
   Grd.Col = 0
   Grd.SetFocus
   Set RdoRes2 = Nothing
   
End Sub

Private Sub ClearGrid()
      
   Grd.Clear
   Grd.Rows = 2
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .ColAlignment(3) = 0
      .Row = 0
      .Col = 0
      .Text = "Op No"
      .ColWidth(0) = 750
      .Col = 1
      .Text = "Shop"
      .ColWidth(1) = 1500
      .Col = 2
      .Text = "Work Center"
      .ColWidth(2) = 1500
      .Col = 3
      .Text = "Comment"
      .ColWidth(3) = 6000
      .Col = 0
   End With
      
End Sub

Private Function TrimComment(sComment As String)
   Dim n As Integer
   On Error GoTo DiaErr1
   If Len(sComment) > 0 Then
      If Len(sComment) > 20 Then sComment = Left(sComment, 20)
      If Asc(Left(sComment, 1)) = 13 Then
         sComment = Right(sComment, Len(sComment) - 2)
      End If
      If Len(sComment) > 2 Then
         n = InStr(3, sComment, Chr(13))
         If n > 0 Then
            TrimComment = Left$(sComment, n - 1)
         Else
            TrimComment = sComment
         End If
      Else
         TrimComment = sComment
      End If
      If Len(TrimComment) < 20 Then
         TrimComment = TrimComment & (String$(20 - (Len(TrimComment)), " "))
      End If
   Else
      TrimComment = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "trimcomme"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

