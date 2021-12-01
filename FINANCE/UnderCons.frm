VERSION 5.00
Begin VB.Form UnderCons
   BorderStyle = 3 'Fixed Dialog
   Caption = "Under Contruction"
   ClientHeight = 1260
   ClientLeft = 2550
   ClientTop = 1830
   ClientWidth = 3930
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 1260
   ScaleWidth = 3930
   ShowInTaskbar = 0 'False
   Begin VB.Timer Timer1
      Interval = 3000
      Left = 240
      Top = 720
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "OK"
      Height = 375
      Left = 1440
      TabIndex = 2
      Top = 720
      Width = 1215
   End
   Begin VB.Label Label2
      BackStyle = 0 'Transparent
      Caption = "We Appologize For Any Incovenience....."
      Height = 255
      Left = 720
      TabIndex = 1
      Top = 360
      Width = 3015
   End
   Begin VB.Label Label1
      BackStyle = 0 'Transparent
      Caption = "This Site Is Currently Being Revised"
      Height = 255
      Left = 720
      TabIndex = 0
      Top = 120
      Width = 3135
   End
   Begin VB.Image Image1
      Height = 480
      Left = 120
      Picture = "UnderCons.frx":0000
      Top = 120
      Width = 480
   End
End
Attribute VB_Name = "UnderCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Activate()
   Beep
   
End Sub

Private Sub Form_Load()
   Move MdiSect.Left + 2500, MdiSect.Top + 1500
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set UnderCons = Nothing
   
End Sub


Private Sub Timer1_Timer()
   Unload Me
   
End Sub
