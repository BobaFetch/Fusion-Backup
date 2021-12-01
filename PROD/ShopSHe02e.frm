VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHe02e 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Manufacturing Order Comments"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe02e.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDmy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "ShopSHe02e.frx":07AE
      DownPicture     =   "ShopSHe02e.frx":1120
      Height          =   350
      Left            =   960
      Picture         =   "ShopSHe02e.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Standard Comments"
      Top             =   1200
      Width           =   350
   End
   Begin VB.TextBox txtCmt 
      Height          =   1455
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2985
      FormDesignWidth =   6885
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   645
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   5
      Top             =   645
      Width           =   615
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   645
      Width           =   375
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   645
      Width           =   3075
   End
End
Attribute VB_Name = "ShopSHe02e"
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
Dim sPartNumber As String

Private Sub cmdCan_Click()
   txtCmt_LostFocus
   Form_Deactivate
   
End Sub


Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 8
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub Form_Activate()
   cmdComments.Enabled = True
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Initialize()
   Move 400, 600
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   lblPrt = ShopSHe02a.cmbPrt
   lblRun = ShopSHe02a.cmbRun
   sPartNumber = Compress(lblPrt)
   GetComment
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'On Error Resume Next
   sSql = "UPDATE RunsTable SET RUNCOMMENTS='" & Replace(txtCmt, "'", "''") & "' " _
          & "WHERE RUNREF='" & sPartNumber & "' AND RUNNO=" & Val(lblRun) & " "
   clsADOCon.ExecuteSQL sSql
   ShopSHe02a.OptCmt.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set ShopSHe02e = Nothing
   
End Sub































Private Sub txtCmt_LostFocus()
   'On Error Resume Next
   'Dim b As Byte
   'b = CheckValidColumn(txtCmt)
   txtCmt = CheckLen(txtCmt, 1020)
   
End Sub



Private Sub GetComment()
   Dim RdoRun As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNCOMMENTS FROM RunsTable WHERE RUNREF='" _
          & sPartNumber & "' AND RUNNO=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_KEYSET)
   If bSqlRows Then
      With RdoRun
         txtCmt = "" & Trim(!RUNCOMMENTS)
         ClearResultSet RdoRun
      End With
   End If
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcomment"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

