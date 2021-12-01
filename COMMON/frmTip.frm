VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClose 
      BackColor       =   &H80000016&
      Height          =   255
      Left            =   5280
      MaskColor       =   &H00C0FFFF&
      Picture         =   "frmTip.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close "
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "Label1"
      Height          =   1335
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Hold down your mouse button and drag this tip wherever you need it"
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "User32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Public Sub MoveForm(btn As Integer)
    Dim lngReturnValue As Long
         If btn = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(frmTip.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
End Sub


Public Sub ShowMoreInfo(Info As String)
    Label1.Caption = Info
    Label1.Refresh
    DoEvents
    
    Me.Width = Label1.Width + 300
    Me.Height = Label1.Height + 560
'    Me.Show (vbModeless)
    btnClose.Top = 0
    btnClose.Left = Me.Width - 285

Me.Show vbModeless
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    'Unload Me
    
End Sub

Private Sub Form_Load()
    Label1.AutoSize = True
    Label1.Left = 30
    Label1.Top = 240
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm (Button)
End Sub

Private Sub Label1_Click()
    'Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm (Button)
End Sub
