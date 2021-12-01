VERSION 5.00
Begin VB.Form InfoShow
   BackColor = &H80000018&
   BorderStyle = 1 'Fixed Single
   ClientHeight = 648
   ClientLeft = 2760
   ClientTop = 3720
   ClientWidth = 3060
   ControlBox = 0 'False
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 648
   ScaleWidth = 3060
   Begin VB.Timer Timer1
      Interval = 3000
      Left = 2520
      Top = 360
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      Height = 495
      Left = 120
      TabIndex = 0
      Top = 120
      Width = 2775
   End
End
Attribute VB_Name = "InfoShow"
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
   '
End Sub


Private Sub Form_Unload(Cancel As Integer)
   WindowState = 1
   Set InfoShow = Nothing
   
   
End Sub


Private Sub lblDsc_Click()
   Unload Me
   
End Sub

Private Sub Timer1_Timer()
   Unload Me
   
End Sub
