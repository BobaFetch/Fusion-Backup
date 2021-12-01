VERSION 5.00
Begin VB.Form RoutRTe01d
   BorderStyle = 3 'Fixed Dialog
   Caption = "Estimated Run Quantity"
   ClientHeight = 2292
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 3696
   ForeColor = &H8000000F&
   Icon = "RoutRTe01d.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 2292
   ScaleWidth = 3696
   ShowInTaskbar = 0 'False
   StartUpPosition = 1 'CenterOwner
   Begin VB.CommandButton cmdCan
      Caption = "Close"
      Height = 435
      Left = 2760
      TabIndex = 1
      Top = 0
      Width = 875
   End
   Begin VB.TextBox txtQty
      Height = 285
      Left = 1800
      TabIndex = 0
      Tag = "1"
      Top = 1560
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Work Center Default Overhead Percentage"
      Height = 255
      Index = 3
      Left = 120
      TabIndex = 6
      Top = 1200
      Width = 3375
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Work Center Standard Labor Rates And"
      Height = 255
      Index = 4
      Left = 120
      TabIndex = 5
      Top = 960
      Width = 3375
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Estimated Parts List Labor Costs Based On"
      Height = 255
      Index = 2
      Left = 120
      TabIndex = 4
      Top = 720
      Width = 3375
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Uses The Entered Quantity To Calculate"
      Height = 255
      Index = 1
      Left = 120
      TabIndex = 3
      Top = 480
      Width = 3375
   End
   Begin VB.Label z1
      Caption = "Enter Run Quantity"
      Height = 255
      Index = 0
      Left = 120
      TabIndex = 2
      Top = 1560
      Width = 1575
   End
End
Attribute VB_Name = "RoutRTe01d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub E_Click()
   
End Sub

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   FormatControls
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   RoutRTe01a.lblRunQty = txtQty
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RoutRTe01d = Nothing
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = Format(Abs(Val(txtQty)), "######0")
   
End Sub
