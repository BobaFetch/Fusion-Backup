VERSION 5.00
Begin VB.Form ActBalance
   BorderStyle = 3 'Fixed Dialog
   ClientHeight = 915
   ClientLeft = 45
   ClientTop = 45
   ClientWidth = 4830
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 915
   ScaleWidth = 4830
   ShowInTaskbar = 0 'False
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Current Account Balance"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 2
      Top = 600
      Width = 2385
   End
   Begin VB.Line Line1
      X1 = 0
      X2 = 4800
      Y1 = 480
      Y2 = 480
   End
   Begin VB.Label lblAcctDesc
      Caption = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Height = 255
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 4575
   End
   Begin VB.Label lblBalance
      Alignment = 1 'Right Justify
      Caption = "$$$$$.$$"
      Height = 255
      Left = 2520
      TabIndex = 0
      Top = 600
      Width = 1335
   End
End
Attribute VB_Name = "ActBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte


Private Sub cmdClose_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
   
   If bOnLoad Then
      bOnLoad = 0
   End If
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
   Move MdiSect.Left + 5000, MdiSect.Top + 1000
   bOnLoad = 1
   Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set ActBalance = Nothing
   
End Sub
