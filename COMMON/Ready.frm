VERSION 5.00
Begin VB.Form Ready
   Appearance = 0 'Flat
   BackColor = &H00C0C0C0&
   BorderStyle = 3 'Fixed Dialog
   Caption = "    "
   ClientHeight = 660
   ClientLeft = 3060
   ClientTop = 2175
   ClientWidth = 3435
   ControlBox = 0 'False
   BeginProperty Font
   Name = "MS Sans Serif"
   Size = 8.25
   Charset = 0
   Weight = 700
   Underline = 0 'False
   Italic = 0 'False
   Strikethrough = 0 'False
   EndProperty
   ForeColor = &H00C0C0C0&
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   MousePointer = 11 'Hourglass
   NegotiateMenus = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 660
   ScaleWidth = 3435
   ShowInTaskbar = 0 'False
   Begin VB.TextBox txtDef
      Appearance = 0 'Flat
      BackColor = &H00C0C0C0&
      BorderStyle = 0 'None
      ForeColor = &H00C0C0C0&
      Height = 285
      Left = 600
      TabIndex = 0
      Top = 600
      Visible = 0 'False
      Width = 345
   End
   Begin VB.Timer tmr1
      Interval = 100
      Left = 2760
      Top = 480
   End
   Begin VB.Label msg
      Alignment = 2 'Center
      BorderStyle = 1 'Fixed Single
      Caption = " "
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 285
      Left = 720
      TabIndex = 1
      Top = 120
      Width = 2475
   End
   Begin VB.Image img1
      Appearance = 0 'Flat
      Height = 480
      Left = 105
      Picture = "Ready.frx":0000
      Top = 40
      Width = 480
   End
End
Attribute VB_Name = "Ready"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Private Sub Form_Load()
   Dim i As Integer
   i = MdiSect.Height - (Height + 500)
   If Screen.Width < 9900 Then
      Move MdiSect.Left + 2100, i
   Else
      If Screen.Width > 13000 Then
         Move MdiSect.Left + 3500, i
      Else
         Move MdiSect.Left + 2700, i
      End If
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   MouseCursor 0
   Set Ready = Nothing
   
End Sub

Private Sub Form_Activate()
   MouseCursor 13
   
End Sub

Private Sub tmr1_Timer()
   Unload Me
   
End Sub
