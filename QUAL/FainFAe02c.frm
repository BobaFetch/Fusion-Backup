VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form FainFAe02c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First Article Item Comments"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCmt 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "9"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   1
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
      FormDesignHeight=   2760
      FormDesignWidth =   5865
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description/Inspection Comments"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2265
   End
   Begin VB.Label txtRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Revision"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label txtPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   1545
   End
End
Attribute VB_Name = "FainFAe02c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim RdoCmt As ADODB.Recordset
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   With RdoCmt
      !FA_ITCOMMENTS = Trim(txtCmt)
      .Update
   End With
   Sleep 50
   Unload Me
   
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      GetFaComment
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set RdoCmt = Nothing
   Set FainFAe02c = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1020)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub



Private Sub GetFaComment()
   On Error GoTo DiaErr1
   sSql = "SELECT FA_ITNUMBER,FA_ITREVISION,FA_ITFEATURENUM," _
          & "FA_ITCOMMENTS FROM FaitTable WHERE (FA_ITNUMBER='" _
          & Compress(txtPrt) & "' AND FA_ITREVISION='" _
          & Trim(txtRev) & "' AND FA_ITFEATURENUM=" & Trim(lblItem) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmt, ES_KEYSET)
   If bSqlRows Then txtCmt = "" & Trim(RdoCmt!FA_ITCOMMENTS)
   Exit Sub
   
DiaErr1:
   sProcName = "getfacomment"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
   DoModuleErrors Me
   
End Sub
