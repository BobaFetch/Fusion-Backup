VERSION 5.00
Begin VB.Form PopMsg
   Appearance = 0 'Flat
   BorderStyle = 1 'Fixed Single
   Caption = "Message."
   ClientHeight = 1065
   ClientLeft = 3060
   ClientTop = 2175
   ClientWidth = 3150
   ControlBox = 0 'False
   FillColor = &H00FFFFFF&
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
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 1065
   ScaleWidth = 3150
   ShowInTaskbar = 0 'False
   StartUpPosition = 1 'CenterOwner
   Begin VB.TextBox txtDef
      Appearance = 0 'Flat
      BackColor = &H8000000F&
      BorderStyle = 0 'None
      ForeColor = &H00C0C0C0&
      Height = 285
      Left = 2730
      TabIndex = 0
      Top = 0
      Visible = 0 'False
      Width = 345
   End
   Begin VB.Timer tmr1
      Interval = 2000
      Left = 0
      Top = 945
   End
   Begin VB.Label lblGreat
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      ForeColor = &H00800000&
      Height = 315
      Left = 600
      TabIndex = 2
      Top = 120
      Visible = 0 'False
      Width = 1935
      WordWrap = -1 'True
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
      Left = 270
      TabIndex = 1
      Top = 540
      Width = 2715
   End
   Begin VB.Image img1
      Appearance = 0 'Flat
      Height = 480
      Left = 105
      Picture = "Popmsg.frx":0000
      Top = 40
      Width = 480
   End
End
Attribute VB_Name = "PopMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim sYear As String
   sYear = Format$(Now, "yyyy")
   Caption = "ES/" & sYear & " ERP Message."
   BackColor = Es_FormBackColor
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Hide
   
End Sub

Private Sub Form_Resize()
   Refresh
   If WindowState = vbMaximized Then
      WindowState = vbNormal
      tmr1.Interval = 10000
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   MouseCursor 0
   Set PopMsg = Nothing
   
End Sub

Private Sub Form_Activate()
   MouseCursor 13
   
End Sub

Private Sub tmr1_Timer()
   Unload Me
   
End Sub
