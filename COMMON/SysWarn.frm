VERSION 5.00
Begin VB.Form SysWarn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ES/2000 ERP"
   ClientHeight    =   1290
   ClientLeft      =   3630
   ClientTop       =   2655
   ClientWidth     =   3255
   ControlBox      =   0   'False
   Icon            =   "SysWarn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "SysWarn.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This Program Requires ES/2000 ERP"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "SysWarn"
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




Private Sub Form_Load()
   SetWindowPos hwnd, hWnd_TopMost, 0, 0, 0, 0, Swp_NOMOVE + Swp_NOSIZE
   Label1 = "This Program Requires ES/2000 ERP"
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SetWindowPos hwnd, Hwnd_NOTOPMOST, 0, 0, 0, 0, Swp_NOMOVE + Swp_NOSIZE
   
End Sub


Private Sub Form_Resize()
   On Error Resume Next
   If WindowState <> 0 Then WindowState = 0
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   MouseCursor 0
   'Set RdoEnv = Nothing
   Set SysWarn = Nothing
End

End Sub
