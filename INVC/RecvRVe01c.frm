VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RecvRVe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PO Item Comments"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RecvRVe01c.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdComments 
      DownPicture     =   "RecvRVe01c.frx":07AE
      Height          =   320
      Index           =   6
      Left            =   5280
      Picture         =   "RecvRVe01c.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Edit PO Item Comments"
      Top             =   840
      Width           =   350
   End
   Begin VB.TextBox txtCmt 
      Height          =   1452
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "9"
      ToolTipText     =   "(2048) Characters Maximum"
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4800
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   5745
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   580
   End
   Begin VB.Label lblPon 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "RecvRVe01c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   '*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
   '*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
   '*** and is protected under US and International copyright    ***
   '*** laws and treaties.                                       ***
   'See the UpdateTables procedure for database revisions
   '4/48/06 Added Comment selection, limited width of data (2048)
   Option Explicit
   Dim bOnLoad As Byte
   
   Private txtKeyPress() As New EsiKeyBd
   Private txtGotFocus() As New EsiKeyBd
   Private txtKeyDown() As New EsiKeyBd
   
   
   Private Sub cmdCan_Click()
      Unload Me
      
   End Sub
   
   Private Sub cmdCmt_Click(Index As Integer)
      
   End Sub
   
   Private Sub cmdComments_Click(Index As Integer)
      If cmdComments(Index) Then
         'See List For Index
         txtCmt.SetFocus
         SysComments.lblListIndex = 1
         SysComments.Show
         cmdComments(Index) = False
      End If
      
   End Sub
   
   
   Private Sub cmdHlp_Click()
      If cmdHlp Then
         MouseCursor 13
         ' SelectHelpTopic Me, "No Subject Help"
         cmdHlp = False
         MouseCursor 0
      End If
      
   End Sub
   
   
   Private Sub Form_Activate()
      If bOnLoad Then GetItemComment
      bOnLoad = 0
      MouseCursor 0
      
   End Sub
   
   Private Sub Form_Deactivate()
      Unload Me
      
   End Sub
   
   
   Private Sub Form_Load()
      FormLoad Me, ES_DONTLIST
      ' MM Move RecvRVe01b.Left + 600, RecvRVe01b.Top + 750
      FormatControls
      bOnLoad = 1
      
   End Sub
   
   
   Private Sub Form_Resize()
      Refresh
      
   End Sub
   
   
   Private Sub Form_Unload(Cancel As Integer)
      Set RecvRVe01c = Nothing
      
   End Sub
   
   
   
   Private Sub FormatControls()
      Dim b As Byte
      b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
      
   End Sub
   
   
   Private Sub GetItemComment()
      Dim RdoCmt As ADODB.Recordset
      On Error GoTo DiaErr1
      sSql = "SELECT PINUMBER,PIITEM,PIREV,PICOMT FROM " _
             & "PoitTable WHERE (PINUMBER=" & Val(lblPon) & " " _
             & "AND PIITEM=" & Val(lblItem) & " AND PIREV='" _
             & lblRev & "')"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmt, ES_STATIC)
      If bSqlRows Then
         With RdoCmt
            txtCmt = "" & Trim(!PICOMT)
            ClearResultSet RdoCmt
         End With
      End If
      Set RdoCmt = Nothing
      Exit Sub
      
DiaErr1:
      CurrError.Number = Err
      CurrError.Description = Err.Description
      DoModuleErrors Me
      
   End Sub
   
   Private Sub txtCmt_LostFocus()
      txtCmt = CheckLen(txtCmt, 2048)
      On Error Resume Next
      txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
      sSql = "UPDATE PoitTable SET PICOMT='" & txtCmt & "' " _
             & "WHERE (PINUMBER=" & Val(lblPon) & " " _
             & "AND PIITEM=" & Val(lblItem) & " AND PIREV='" _
             & lblRev & "')"
      clsADOCon.ExecuteSQL sSql
      
   End Sub
