VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PurcPRe05b 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Buyers To Vendors"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "PurcPRe05b.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdVnd 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "Assign Vendors To This Buyer"
      Top             =   120
      Width           =   875
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Click Vendor Or Use SpaceBar To Select - ESC To Cancel"
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   2280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3960
      Width           =   875
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   3840
      Picture         =   "PurcPRe05b.frx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   4200
      Picture         =   "PurcPRe05b.frx":0694
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblByr 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer ID"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "PurcPRe05b"
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


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdVnd_Click()
   Dim bResponse As Byte
   Dim iList As Integer
   Dim sMsg As String
   Dim sBuyer As String
   Dim sVendor As String
   
   sBuyer = Compress(lblByr)
   sMsg = "Do You Wish To Update The Vendors To The Current Buyer?" _
          & vbCr & "Note Only The Vendor Record Is Changed. No Others."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   On Error Resume Next
   If bResponse = vbYes Then
      For iList = 1 To Grid1.Rows - 1
         Grid1.row = iList
         Grid1.Col = 0
         sVendor = Compress(Grid1.Text)
         Grid1.Col = 2
         If Grid1.CellPicture = Chkyes.Picture Then
            sSql = "UPDATE VndrTable SET VEBUYER='" & sBuyer _
                   & "' WHERE VEREF='" & sVendor & "'"
         Else
            sSql = "UPDATE VndrTable SET VEBUYER='' WHERE " _
                   & "VEREF='" & sVendor & "'"
         End If
         clsADOCon.ExecuteSQL sSql
      Next
      MsgBox "Vendors Have Been Updated.", _
         vbInformation, Caption
      Unload Me
   Else
      CancelTrans
   End If
   
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillGrid
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   On Error Resume Next
   If MDISect.SideBar.Visible = False Then
      Move MDISect.Left + MDISect.ActiveForm.Left + 200, MDISect.Top + 2200
   Else
      Move MDISect.Left + MDISect.ActiveForm.Left + 1800, MDISect.Top + 2800
   End If
   bOnLoad = 1
   With Grid1
      .Rows = 2
      .ColWidth(0) = 700
      .ColWidth(1) = 1450
      .ColWidth(2) = 2850
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .ColAlignment(2) = 0
      .row = 0
      .Col = 0
      .Text = "Assigned"
      .Col = 1
      .Text = "Vendor "
      .Col = 2
      .Text = "Name"
   End With
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   PurcPRe05a.optVew.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set PurcPRe05b = Nothing
   
End Sub





Private Sub FillGrid()
   Dim RdoVdr As ADODB.Recordset
   Dim iList As Integer
   Dim sBuyer As String
   On Error Resume Next
   Grid1.Rows = 1
   Grid1.row = 1
   
   On Error Resume Next
   sSql = "SELECT VEREF,VENICKNAME,VEBNAME,VEBUYER FROM VndrTable " _
          & "WHERE VEREF<>'NONE'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVdr, ES_FORWARD)
   If bSqlRows Then
      With RdoVdr
         sBuyer = Compress(lblByr)
         Do Until .EOF
            iList = iList + 1
            Grid1.Rows = iList + 1
            Grid1.Col = 1
            Grid1.row = iList
            Grid1.Text = "" & Trim(!VENICKNAME)
            Grid1.Col = 2
            Grid1.Text = "" & Trim(!VEBNAME)
            Grid1.Col = 0
            If Trim(!VEBUYER) = sBuyer Then
               Set Grid1.CellPicture = Chkyes.Picture
            Else
               Set Grid1.CellPicture = Chkno.Picture
            End If
            .MoveNext
         Loop
         ClearResultSet RdoVdr
      End With
      If Grid1.Rows > 0 Then cmdVnd.Enabled = True
      Grid1.row = 1
      Grid1.Col = 0
      Grid1.SetFocus
   End If
   Set RdoVdr = Nothing
   
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
      Grid1.Col = 0
      If Grid1.Text = "X" Then
         Set Grid1.CellPicture = Chkyes.Picture
         Grid1.Text = ""
      Else
         Set Grid1.CellPicture = Chkyes.Picture
         Grid1.Text = "X"
      End If
      Grid1.Col = 0
   End If
   
End Sub


Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grid1.Col = 0
   If Grid1.CellPicture = Chkno.Picture Then
      Set Grid1.CellPicture = Chkyes.Picture
   Else
      Set Grid1.CellPicture = Chkno.Picture
   End If
   Grid1.Col = 0
   
End Sub
