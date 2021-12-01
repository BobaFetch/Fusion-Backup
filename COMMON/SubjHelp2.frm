VERSION 5.00
Begin VB.Form SubjHelp
   BackColor = &H80000018&
   BorderStyle = 1 'Fixed Single
   ClientHeight = 2040
   ClientLeft = 12
   ClientTop = 12
   ClientWidth = 2820
   ClipControls = 0 'False
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 2040
   ScaleWidth = 2820
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox optFrom
      Height = 255
      Left = 1680
      TabIndex = 2
      Top = 1800
      Visible = 0 'False
      Width = 495
   End
   Begin VB.Timer Timer1
      Enabled = 0 'False
      Interval = 32000
      Left = 0
      Top = 1680
   End
   Begin VB.Label Section
      BackStyle = 0 'Transparent
      BeginProperty Font
      Name = "MS Sans Serif"
      Size = 7.8
      Charset = 0
      Weight = 400
      Underline = -1 'True
      Italic = 0 'False
      Strikethrough = 0 'False
      EndProperty
      Height = 255
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 2535
   End
   Begin VB.Label hlp
      BackStyle = 0 'Transparent
      Height = 1575
      Left = 120
      TabIndex = 0
      Top = 360
      Width = 2535
   End
End
Attribute VB_Name = "SubjHelp"
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

Private Sub Form_GotFocus()
   Timer1.Enabled = True
   
End Sub


Private Sub Form_Load()
   Show
   
End Sub


Private Sub Label1_Click()
   
End Sub


Private Sub Form_LostFocus()
   Timer1.Enabled = False
   Unload Me
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set SubjHelp = Nothing
   
End Sub

Private Sub hlp_Click()
   Unload Me
   
End Sub


Private Sub Timer1_Timer()
   Unload Me
   
End Sub
