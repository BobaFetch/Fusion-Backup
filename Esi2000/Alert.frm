VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form Alert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  "
   ClientHeight    =   1515
   ClientLeft      =   -4920
   ClientTop       =   225
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   360
      Top             =   0
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   3600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   1200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1515
      FormDesignWidth =   4695
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Alert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of          ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Dim sOldMessage As String

Private Sub cmdCan_Click()
   'move off screen
   Caption = ""
   Move (Width - (Width * 2) - 200), 0
   
End Sub



Private Sub Form_Load()
   sOldMessage = ""
   lblMsg.ForeColor = ES_RED
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set Alert = Nothing
   
End Sub





Public Sub GetMessage()
   Dim RdoMsg As ADODB.Recordset
   Timer1.Enabled = False
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM Alerts WHERE ALERTREF=1 AND ALERTMSG<>''"
   'Set RdoMsg = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   Set RdoMsg = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
   If Not RdoMsg.BOF And Not RdoMsg.EOF Then
      If RdoMsg!ALERTMSG <> "" Then
         Caption = "System Message"
         If sOldMessage <> "" & Trim(RdoMsg!ALERTMSG) Then
            Beep
            AppActivate App.Title
            Move (Screen.Width / 2) - (Width / 2), (Screen.Height / 2) - (Height / 2)
            lblMsg = "" & Trim(RdoMsg!ALERTMSG)
            sOldMessage = lblMsg
         End If
      End If
   End If
   Timer1.Enabled = True
   Exit Sub
DiaErr1:
   Timer1.Enabled = True
   On Error GoTo 0
   
End Sub

Private Sub Timer1_Timer()
   GetMessage
   
End Sub
