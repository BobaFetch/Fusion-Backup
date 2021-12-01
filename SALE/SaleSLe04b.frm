VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SaleSLe04b 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Book Customers"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLe04b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "S&elect"
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Selects A Maximum Of 300 Items"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Pattern Matches Leading Character >=  (Fills Up To 300 Customers)"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      ToolTipText     =   "Update Changes To The List"
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   4680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5235
      FormDesignWidth =   5925
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3615
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Click The Row To Select/Unselect A Customer (300 Customers Max)"
      Top             =   1080
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   8404992
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Label lblSelected 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1515
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   480
      Picture         =   "SaleSLe04b.frx":07AE
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   240
      Picture         =   "SaleSLe04b.frx":0B38
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblPriceBook 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customers For Price Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "SaleSLe04b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/22/04 Added the Selection features
'9/9/06 Revised Price Book trap (UpdateCustomers)
Option Explicit
Dim bOnLoad As Byte
Dim bChanged As Byte

Dim sCustomers(400, 3) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdGo_Click(Index As Integer)
   FillGrid
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   UpdateCustomers
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      lblPriceBook = SaleSLe04a.cmbPrb
      FillGrid
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Move SaleSLe04a.Left + 800, SaleSLe04a.Top + 1200
   BackColor = Es_HelpBackGroundColor
   FormatControls
   With grd
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Listed"
      .Col = 1
      .Text = "Nickname"
      .Col = 2
      .Text = "Name"
      .Col = 3
      .Text = "Current Book"
      .ColWidth(0) = 700
      .ColWidth(1) = 1200
      .ColWidth(2) = 2100
      .ColWidth(3) = 1500
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   If bChanged = 1 Then
      bResponse = MsgBox("Exit Without Saving Changes?", ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   SaleSLe04a.optCst.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SaleSLe04b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Private Sub FillGrid()
   Dim RdoGrd As ADODB.Recordset
   Dim bLength As Byte
   Dim sBook As String
   sBook = Compress(lblPriceBook)
   On Error GoTo DiaErr1
   grd.Rows = 1
   Erase sCustomers
   bLength = Len(Trim(txtCst))
   If bLength = 0 Then bLength = 1
   If bLength < 4 Then
''      sSql = "SELECT CUREF,CUNICKNAME,CUNAME,CUPRICEBOOK," _
''             & "PBHREF,PBHID FROM CustTable,PbhdTable WHERE " _
''             & "(CUPRICEBOOK *=PBHREF AND LEFT(CUREF," & Str$(bLength) & ") " _
''             & ">= '" & Compress(Left$(txtCst, bLength)) & "') ORDER BY CUREF"
'      sSql = "SELECT CUREF,CUNICKNAME,CUNAME,CUPRICEBOOK," _
'             & "PBHREF,PBHID" & vbCrLf _
'             & "FROM CustTable" & vbCrLf _
'             & "LEFT JOIN PbhdTable ON CUPRICEBOOK=PBHREF" & vbCrLf _
'             & "WHERE LEFT(CUREF," & str$(bLength) & ") " _
'             & ">= '" & Compress(Left$(txtCst, bLength)) & "'" & vbCrLf _
'             & "ORDER BY CUREF"
      sSql = "SELECT CUREF,CUNICKNAME,CUNAME,CUPRICEBOOK,PBHREF,PBHID" & vbCrLf _
         & "FROM CustTable" & vbCrLf _
         & "LEFT JOIN PbhdTable ON CUPRICEBOOK=PBHREF" & vbCrLf
      If Len(Compress(txtCst)) > 0 Then
         sSql = sSql & "WHERE CUREF LIKE '" & Compress(txtCst) & "%'" & vbCrLf
      End If
   Else
'      sSql = "SELECT CUREF,CUNICKNAME,CUNAME,CUPRICEBOOK," _
'             & "PBHREF,PBHID FROM CustTable,PbhdTable WHERE " _
'             & "(CUPRICEBOOK *=PBHREF AND CUREF LIKE '" _
'             & Compress(txtCst) & "%') ORDER BY CUREF"
      sSql = "SELECT CUREF,CUNICKNAME,CUNAME,CUPRICEBOOK," _
             & "PBHREF,PBHID" & vbCrLf _
             & "FROM CustTable" & vbCrLf _
             & "LEFT JOIN PbhdTable ON CUPRICEBOOK=PBHREF" & vbCrLf _
             & "WHERE CUREF LIKE '" & Compress(txtCst) & "%'" & vbCrLf
   End If
   sSql = sSql & "ORDER BY CUREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
   If bSqlRows Then
      MouseCursor 13
      With RdoGrd
         Do Until .EOF
            grd.Rows = grd.Rows + 1
            If grd.Rows > 300 Then Exit Do
            grd.Row = grd.Rows - 1
            sCustomers(grd.Row, 0) = "" & Trim(!CUREF)
            grd.Col = 0
            If "" & Trim(!CUPRICEBOOK) = sBook Then
               Set grd.CellPicture = Chkyes.Picture
               sCustomers(grd.Row, 1) = "X"
            Else
               Set grd.CellPicture = Chkno.Picture
               sCustomers(grd.Row, 1) = ""
            End If
            sCustomers(grd.Row, 2) = "" & Trim(!CUPRICEBOOK)
            grd.Col = 1
            grd.Text = "" & Trim(!CUNICKNAME)
            grd.Col = 2
            grd.Text = "" & Trim(!CUNAME)
            grd.Col = 3
            grd.Text = "" & Trim(!PBHID)
            .MoveNext
         Loop
         ClearResultSet RdoGrd
         bChanged = 0
      End With
      grd.Col = 0
      grd.Row = 1
   End If
   lblSelected = "Selected " & Val(grd.Rows - 1) & " Customers"
   MouseCursor 0
   Set RdoGrd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'9/9/06 Revised Price Book trap

Private Sub UpdateCustomers()
   Dim iList As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error Resume Next
   sMsg = "This Procedure Will Add Or Remove Customers" & vbCrLf _
          & "From This Price Book.  Continue?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      For iList = 1 To grd.Rows - 1
         If sCustomers(iList, 1) = "X" Then
            sSql = "UPDATE CustTable SET CUPRICEBOOK='" _
                   & Compress(lblPriceBook) & "' WHERE CUREF='" _
                   & sCustomers(iList, 0) & "' "
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
         Else
            If sCustomers(iList, 2) = Compress(lblPriceBook) Then
               sSql = "UPDATE CustTable SET CUPRICEBOOK='' " _
                      & "WHERE CUREF='" & sCustomers(iList, 0) & "' " _
                      & "AND CUPRICEBOOK='" & Compress(lblPriceBook) & "'"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
            End If
         End If
      Next
      MouseCursor 0
      SysMsg "Changes Were Saved.", True
      bChanged = 0
      Unload Me
   Else
      CancelTrans
   End If
   
End Sub



Private Sub grd_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      grd.Col = 0
      If grd.CellPicture = Chkyes.Picture Then
         Set grd.CellPicture = Chkno.Picture
         sCustomers(grd.Row, 1) = ""
      Else
         Set grd.CellPicture = Chkyes.Picture
         sCustomers(grd.Row, 1) = "X"
      End If
      bChanged = 1
   End If
   
End Sub

Private Sub grd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   On Error Resume Next
   grd.Col = 0
   If grd.CellPicture = Chkyes.Picture Then
      Set grd.CellPicture = Chkno.Picture
      sCustomers(grd.Row, 1) = ""
   Else
      Set grd.CellPicture = Chkyes.Picture
      sCustomers(grd.Row, 1) = "X"
   End If
   bChanged = 1
   
End Sub

Private Sub txtCst_LostFocus()
   txtCst = CheckLen(txtCst, 10)
   
End Sub
