VERSION 5.00
Begin VB.Form Awarn
   BorderStyle = 3 'Fixed Dialog
   Caption = "ES/2000 ERP"
   ClientHeight = 1380
   ClientLeft = 3630
   ClientTop = 2655
   ClientWidth = 3255
   ControlBox = 0 'False
   Icon = "diaEnd.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 1380
   ScaleWidth = 3255
   ShowInTaskbar = 0 'False
   StartUpPosition = 2 'CenterScreen
   Begin VB.PictureBox Picture1
      BorderStyle = 0 'None
      Height = 495
      Left = 120
      Picture = "diaEnd.frx":030A
      ScaleHeight = 495
      ScaleWidth = 495
      TabIndex = 2
      Top = 0
      Width = 495
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "OK"
      Height = 375
      Left = 1320
      TabIndex = 0
      Top = 840
      Width = 855
   End
   Begin VB.Label Label1
      Alignment = 2 'Center
      Caption = "This Program Requires ES/2000 ERP"
      Height = 255
      Left = 240
      TabIndex = 1
      Top = 480
      Width = 2775
   End
End
Attribute VB_Name = "Awarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Dim sYear As String
   sYear = "ES/" & Format$(Now, "yyyy") & " ERP"
   Caption = sYear
   Label1 = "This Program Requires " & sYear
   Y = False
   
End Sub


Private Sub Form_Resize()
   On Error Resume Next
   If WindowState <> 0 Then WindowState = 0
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Y = True
   Set Awarn = Nothing
   
End Sub
